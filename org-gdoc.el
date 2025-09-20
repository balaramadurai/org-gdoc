;;; org-gdoc.el --- Convert local .gddoc files to Org-mode  -*- lexical-binding: t; -*-

;; Copyright (C) 2025 Bala Ramadurai <bala@balaramadurai.net>

;; Author: Bala Ramadurai <bala@balaramadurai.net>
;; Version: 0.1
;; Package-Requires: ((emacs "24.4"))
;; Keywords: org-mode, google-docs, conversion
;; URL: https://github.com/balaramadurai/org-gdoc
;; License: MIT

;; This file is not part of GNU Emacs.

;; This program is free software: you can redistribute it and/or modify
;; it under the terms of the MIT License. See the LICENSE file for
;; details.

;;; Commentary:
;; This package allows conversion of locally synced Google Docs (.gddoc files)
;; from Insync to Org-mode files in Emacs. It supports converting a single file
;; or a directory of files, with an option to merge all into one Org file.

;;; Code:

(defcustom org-gdoc-python-script-path "~/gdoc_to_org.py"
  "Path to the Python script for converting Google Docs."
  :type 'string
  :group 'org-gdoc)

(defcustom org-gdoc-credentials-path "~/Documents/3Resources/emacs/pure-emacs/gdoc-api-credentials.json"
  "Path to the Google API credentials JSON file."
  :type 'string
  :group 'org-gdoc)

(defun org-gdoc-extract-metadata (gddoc-file)
  "Extract the Google Doc metadata (ID and URL) from a .gddoc file."
  (with-temp-buffer
    (insert-file-contents gddoc-file)
    (let ((json-object-type 'hash-table)
          (json-array-type 'list)
          (json-key-type 'string))
      (condition-case err
          (let ((json-data (json-read)))
            (list (gethash "file_id" json-data)
                  (gethash "url" json-data)))
        (error (message "Error parsing %s: %s" gddoc-file err)
               (list nil "N/A"))))))

(defun org-gdoc-convert (gddoc-file)
  "Convert the Google Doc from GDDOC-FILE and return the Org-mode content."
  (let* ((metadata (org-gdoc-extract-metadata gddoc-file))
         (doc-id (car metadata))
         (python-bin (executable-find "python3"))
         (convert-command (format "%s %s %s %s"
                                  python-bin
                                  (shell-quote-argument (expand-file-name org-gdoc-python-script-path))
                                  (shell-quote-argument doc-id)
                                  (shell-quote-argument (expand-file-name org-gdoc-credentials-path))))
         (output (if doc-id (shell-command-to-string convert-command) "")))
    (if (string-match "^Error: " output)
        (progn (message "Conversion error for %s: %s" gddoc-file output) "")
      (let ((lines (split-string output "\n" t)))
        (if (and lines (string-match "^\\* [0-9a-zA-Z_-]\\{44\\}$" (car lines))) ; Match title pattern (e.g., * file_id)
            (string-join (cdr lines) "\n")
          output))))) ; Keep all lines if no title pattern

(defun org-gdoc-convert-directory (output-file directory &optional merge)
  "Convert all .gddoc files in DIRECTORY to Org-mode.
If MERGE is non-nil, merge into OUTPUT-FILE with headings and drawers; otherwise, create individual files."
  (let ((gddoc-files (directory-files directory t "\\.gddoc$")))
    (if merge
        (let ((merged-content ""))
          (dolist (file gddoc-files)
            (let* ((metadata (org-gdoc-extract-metadata file))
                   (doc-id (car metadata))
                   (url (cadr metadata))
                   (title (file-name-sans-extension (file-name-nondirectory file)))
                   (content (org-gdoc-convert file)))
              (when doc-id
                (setq merged-content (concat merged-content "\n* " title "\n:PROPERTIES:\n:URL: " url "\n:END:\n" content)))))
          (with-temp-file output-file
            (insert (substring merged-content 1))) ; Remove leading newline
          (message "Merged %d Google Docs into %s" (length gddoc-files) output-file))
      (dolist (file gddoc-files)
        (let* ((base-name (file-name-sans-extension (file-name-nondirectory file)))
               (org-file (expand-file-name (concat base-name ".org") (file-name-directory file))))
          (with-temp-file org-file
            (insert (org-gdoc-convert file)))
          (message "Converted %s to %s" (file-name-nondirectory file) org-file))))))

(defun org-gdoc-download (output-file)
  "Convert .gddoc files to Org-mode.
Prompt for a single .gddoc file or a directory, with an option to merge into one Org file."
  (interactive "FOutput Org-mode file or directory: ")
  (let ((choice (read-char-choice "Convert [f]ile, [d]irectory, or [m]erge directory? " '(?f ?d ?m))))
    (cond
     ((eq choice ?f)
      (let ((gddoc-file (read-file-name "Select .gddoc file: " "~/" nil t ".gddoc")))
        (if (file-exists-p gddoc-file)
            (with-temp-file output-file
              (insert (org-gdoc-convert gddoc-file)))
          (message "Invalid .gddoc file selected"))))
     ((eq choice ?d)
      (let ((directory (read-directory-name "Select directory containing .gddoc files: " "~/Insync/bala@balaramadurai.net/" nil t)))
        (if (file-directory-p directory)
            (org-gdoc-convert-directory output-file directory nil)
          (message "Invalid directory selected"))))
     ((eq choice ?m)
      (let ((directory (read-directory-name "Select directory containing .gddoc files to merge: " "~/Insync/bala@balaramadurai.net/" nil t)))
        (if (file-directory-p directory)
            (org-gdoc-convert-directory output-file directory t)
          (message "Invalid directory selected"))))
     (t (message "Invalid choice")))))

(defun org-gdoc-push-suggestion ()
  "Push the current Org subtree as a tracked suggestion to the linked Google Doc."
  (interactive)
  (let ((url (org-entry-get nil "URL")))
    (if url
        (let* ((doc-id (substring url (string-match "/d/\\([^/]+\\)" url) (match-end 1)))
               (subtree-content (buffer-substring-no-properties (org-entry-beginning-position) (org-entry-end-position)))
               (python-bin (executable-find "python3"))
               (push-command (format "%s %s %s %s --push-suggestion %s"
                                     python-bin
                                     (shell-quote-argument (expand-file-name org-gdoc-python-script-path))
                                     (shell-quote-argument doc-id)
                                     (shell-quote-argument (expand-file-name org-gdoc-credentials-path))
                                     (shell-quote-argument subtree-content)))
               (output (shell-command-to-string push-command)))
          (if (string-match "^Error: " output)
              (message "Push failed: %s" output)
            (message "Pushed suggestion to Google Doc")))
      (message "No URL found in subtree properties."))))

(provide 'org-gdoc)

;;; org-gdoc.el ends here
