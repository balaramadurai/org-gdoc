;;; org-gdoc.el --- Sync a single Org file with a folder of Google Docs.
;;
;; Copyright (C) 2025 Bala Ramadurai <bala@balaramadurai.net>
;;
;; Author: Bala Ramadurai <bala@balaramadurai.net>
;; Keywords: org, gdoc, sync, pandoc
;; URL: https://github.com/balaramadurai/org-gdoc
;; Package-Requires: ((emacs "26.1"))
;; Version: 1.0
;;
;; This file is not part of GNU Emacs.
;;
;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the MIT License.
;;
;; See the LICENSE file for details.

;;; Commentary:
;; This package provides Emacs Lisp functions to call the `org_gdoc_sync.py`
;; Python script. It allows for importing a folder of .docx files into a
;; single org file, and for exporting subtrees or the entire buffer back
;; to their respective .docx files.
;;
;; This script was developed with significant assistance from Google's
;; Gemini 2.5 Pro model.

;;; Code:

(defgroup org-gdoc nil
  "Settings for org-gdoc synchronization."
  :group 'org)

(defcustom org-gdoc-python-executable "python3"
  "The Python executable to use for running the sync script."
  :type 'string
  :group 'org-gdoc)

(defcustom org-gdoc-python-script
  (expand-file-name "org_gdoc_sync.py"
                    (file-name-directory (or load-file-name buffer-file-name)))
  "Path to the org_gdoc_sync.py script.
Assumes it is in the same directory as this .el file by default."
  :type 'string
  :group 'org-gdoc)

(defcustom org-gdoc-reference-docx-path nil
  "Path to a .docx file to use as a style reference for exports.
Set this to a .docx file with your preferred styles (e.g., fonts,
headings) to maintain consistent formatting. If nil, pandoc's
default styles are used."
  :type '(file :must-match t)
  :group 'org-gdoc)

;;;###autoload
(defun org-gdoc-import-from-folder ()
  "Import all .docx files from a directory into a new Org buffer."
  (interactive)
  (let* ((docx-dir (read-directory-name "Import .docx from directory: "))
         (org-file (read-file-name "Save imported Org file as: " (file-name-as-directory docx-dir) nil nil nil)))
    (unless (string-match-p "\\.org$" org-file)
      (setq org-file (concat org-file ".org")))

    (message "Importing from %s..." docx-dir)
    (let ((command (format "%s %s import --folder %s --output %s"
                           (shell-quote-argument org-gdoc-python-executable)
                           (shell-quote-argument org-gdoc-python-script)
                           (shell-quote-argument docx-dir)
                           (shell-quote-argument org-file))))
      (shell-command command)
      (find-file org-file)
      (message "Successfully imported documents into %s" org-file))))

(defun org-gdoc--export-subtree-at (point)
  "Internal function to export the subtree at a given POINT."
  (save-excursion
    (goto-char point)
    (org-back-to-heading t)
    (let* ((source-file (org-entry-get nil "SOURCE_FILE"))
           (beg (point))
           (end (save-excursion (org-end-of-subtree t) (point)))
           (sync-buffer-name "*org-gdoc-sync*"))
      (if (and source-file (not (string-empty-p source-file)))
          (progn
            (message "Syncing subtree to %s..." source-file)
            (let* ((process-args
                    (append (list org-gdoc-python-script
                                  "export"
                                  "--quiet"
                                  "--target-file"
                                  source-file)
                            (when (and org-gdoc-reference-docx-path (file-exists-p org-gdoc-reference-docx-path))
                              (list "--reference-doc" org-gdoc-reference-docx-path)))))
              (let ((exit-code (apply #'call-process-region beg end org-gdoc-python-executable nil `(t ,sync-buffer-name) nil process-args)))
                (if (= exit-code 0)
                    (progn
                      (message "Subtree successfully synced to %s" source-file)
                      ;; On success, we don't need the output buffer.
                      (when (get-buffer sync-buffer-name)
                        (kill-buffer sync-buffer-name)))
                  (progn
                    (pop-to-buffer sync-buffer-name)
                    (error "Org-gdoc sync failed. See %s buffer for details." sync-buffer-name))))))
        (warn "No :SOURCE_FILE: property found in the current subtree. Skipping.")))))

;;;###autoload
(defun org-gdoc-export-subtree ()
  "Export the current Org subtree back to its original .docx file."
  (interactive)
  (unless (derived-mode-p 'org-mode)
    (error "Not in an Org-mode buffer."))
  (org-gdoc--export-subtree-at (point)))

;;;###autoload
(defun org-gdoc-export-buffer ()
  "Export all subtrees in the current buffer to their respective .docx files."
  (interactive)
  (unless (derived-mode-p 'org-mode)
    (error "Not in an Org-mode buffer."))
  (let ((count 0))
    ;; --- FIX: Use `org-map-entries` for a robust iteration ---
    ;; This is the canonical way to iterate over specific headings in Org mode.
    ;; It avoids manual point management and potential infinite loops.
    ;; We match all headings at LEVEL 1, as requested.
    (org-map-entries
     (lambda ()
       (org-gdoc--export-subtree-at (point))
       (setq count (1+ count)))
     "LEVEL=1"
     'file)
    (message "Finished syncing buffer. Processed %d level-1 subtrees." count)))


(provide 'org-gdoc)

;;; org-gdoc.el ends here

