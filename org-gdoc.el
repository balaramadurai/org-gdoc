;;; org-gdoc.el --- Sync between a single Org-mode file and a folder of .docx files -*- lexical-binding: t; -*-

;; Copyright © 2025 Bala Ramadurai <bala@balaramadurai.net>
;;
;; Author: Bala Ramadurai <bala@balaramadurai.net>
;; Maintainer: Bala Ramadurai <bala@balaramadurai.net>
;; Created: 2025-09-10
;; Version: 1.0
;; Package-Requires: ((emacs "27.1"))
;; Homepage: https://github.com/balaramadurai/org-gdoc
;;
;; This file is not part of GNU Emacs.
;;
;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the MIT License.
;;
;; This script was generated with the assistance of Gemini 2.5 Pro.

;;; Commentary:
;;
;; This package provides a bridge between a single Org-mode file used for
;; long-form writing (e.g., a manuscript) and a folder of .docx files that
;; an editor might work with. It allows for importing a directory of .docx
;; files into a single Org file, and exporting changes from Org subtrees
;; back to their corresponding .docx files.

;;; Code:

(defgroup org-gdoc nil
  "Settings for the org-gdoc sync tool."
  :group 'org)

(defcustom org-gdoc-python-executable "python3"
  "The path to the Python executable to run the sync script."
  :type 'string
  :group 'org)

;; Correctly determine the path to the python script relative to this elisp file.
(defvar org-gdoc-python-script
  (let* ((el-file (or load-file-name buffer-file-name))
         (el-dir (file-name-directory el-file)))
    ;; If loaded by straight.el, the file is in a 'build' dir.
    ;; The python script is in the corresponding 'repos' dir.
    (if (string-match-p "/straight/build/" (directory-file-name el-dir))
        (expand-file-name "org_gdoc_sync.py"
                          (replace-regexp-in-string "/build/" "/repos/" el-dir))
      ;; Fallback for local/manual loading
      (expand-file-name "org_gdoc_sync.py" el-dir)))
  "The full path to the org_gdoc_sync.py script.")

(defcustom org-gdoc-reference-docx-path nil
  "Path to a .docx file to use as a style reference for exports.
This is crucial for preserving fonts and formatting. See the
README for instructions on how to create this file properly
using LibreOffice or Microsoft Word."
  :type '(file :must-match t)
  :group 'org-gdoc)

(defun org-gdoc--run-import (args)
  "Internal helper to run the import process and handle errors."
  (let* ((sync-buffer-name "*org-gdoc-sync*")
         (output-buffer (get-buffer-create sync-buffer-name))
         (process-args (append (list org-gdoc-python-script "import") args)))
    (with-current-buffer output-buffer
      (erase-buffer))
    (let* ((exit-code (apply #'call-process org-gdoc-python-executable nil output-buffer nil process-args)))
      (if (= exit-code 0)
          (progn
            (when (get-buffer sync-buffer-name)
              (kill-buffer sync-buffer-name))
            t)
        (progn
          (pop-to-buffer sync-buffer-name)
          (error "Org-gdoc import failed. See the %s buffer for details." sync-buffer-name)
          nil)))))

;;;###autoload
(defun org-gdoc-import-from-folder ()
  "Import all .docx files from a folder into a single Org-mode file."
  (interactive)
  (let* ((folder (read-directory-name "Import from folder: "))
         (target-file (read-file-name "Into Org file: " nil nil t)))
    (message "Importing from %s..." folder)
    (when (org-gdoc--run-import (list "--folder" folder "--output" target-file))
      (message "Successfully imported documents into %s" target-file))))

;;;###autoload
(defun org-gdoc-import-single-file ()
  "Import a single .docx file into an Org-mode file."
  (interactive)
  (let* ((source-file (read-file-name "Import .docx file: " nil nil t))
         (target-file (read-file-name "Into Org file: " nil nil t)))
    (message "Importing %s into %s..." source-file target-file)
    (when (org-gdoc--run-import (list "--file" source-file "--output" target-file))
      (message "Successfully imported %s into %s" source-file target-file))))

;;;###autoload
(defun org-gdoc-insert-subtree-from-new-file ()
  "Prompt for a .docx file and insert it as a new subtree at point."
  (interactive)
  (let ((source-file (read-file-name "Import .docx file as new subtree: ")))
    (if (not (file-exists-p source-file))
        (error "File does not exist: %s" source-file)
      (message "Importing %s as a new subtree..." source-file)
      (let ((temp-file (make-temp-file "org-gdoc-insert-")))
        (unwind-protect
            (progn
              (when (org-gdoc--run-import (list "--file" source-file "--output" temp-file))
                (save-excursion
                  (let ((content (with-temp-buffer
                                   (insert-file-contents temp-file)
                                   (buffer-string))))
                    (insert (string-trim content) "\n")))
                (message "Successfully imported %s as a new subtree." source-file)))
          (when (file-exists-p temp-file)
            (delete-file temp-file)))))))

;;;###autoload
(defun org-gdoc-update-subtree-from-source ()
  "Update current subtree by re-importing from its :SOURCE_FILE:."
  (interactive)
  (let ((source-file (org-entry-get nil "SOURCE_FILE")))
    (if (and source-file (file-exists-p (expand-file-name source-file)))
        (let* ((bounds (org-element-context))
               (start (org-element-property :begin bounds))
               (end (org-element-property :end bounds))
               (temp-file (make-temp-file "org-gdoc-update-")))
          (unwind-protect
              (progn
                (message "Updating subtree from %s..." source-file)
                (when (org-gdoc--run-import (list "--file" source-file "--output" temp-file))
                  (let ((new-content (with-temp-buffer
                                       (insert-file-contents temp-file)
                                       (buffer-string))))
                    (delete-region start end)
                    (goto-char start)
                    (insert (string-trim new-content) "\n")
                    (message "Successfully updated subtree from %s" source-file))))
            (when (file-exists-p temp-file)
              (delete-file temp-file))))
      (warn "No valid :SOURCE_FILE: property found for the current subtree."))))

(defun org-gdoc--export-subtree-at (point)
  "Helper function to export the subtree at a given POINT."
  (goto-char point)
  (let ((source-file (org-entry-get nil "SOURCE_FILE")))
    (if (and source-file (file-exists-p (expand-file-name source-file)))
        (let* ((bounds (org-element-context))
               (start (org-element-property :begin bounds))
               (end (org-element-property :end bounds))
               (args (list org-gdoc-python-script
                           "export"
                           "--target-file" source-file
                           "--quiet"))
               (sync-buffer-name "*org-gdoc-sync*"))
          (when org-gdoc-reference-docx-path
            (setq args (append args (list "--reference-doc" org-gdoc-reference-docx-path))))
          (message "Syncing subtree to %s..." source-file)
          (let ((exit-code (apply #'call-process-region start end org-gdoc-python-executable nil (get-buffer-create sync-buffer-name) nil args)))
            (if (= exit-code 0)
                (progn
                  (when (get-buffer sync-buffer-name)
                    (kill-buffer sync-buffer-name))
                  (message "Subtree successfully synced to %s" source-file))
              (progn
                (pop-to-buffer sync-buffer-name)
                (error "Org-gdoc export failed. See *org-gdoc-sync* buffer for details.")))))
      (warn "No valid :SOURCE_FILE: property found for the current subtree."))))

;;;###autoload
(defun org-gdoc-export-subtree ()
  "Export the current Org subtree back to its original .docx file."
  (interactive)
  (org-gdoc--export-subtree-at (point)))

;;;###autoload
(defun org-gdoc-export-buffer ()
  "Iterate through all level-1 headings and export each to its .docx file."
  (interactive)
  (save-excursion
    (message "Starting full buffer export...")
    (org-map-entries #'org-gdoc--export-subtree-at "LEVEL=1" 'file))
  (message "Full buffer export complete."))

(provide 'org-gdoc)

;;; org-gdoc.el ends here

