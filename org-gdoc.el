;;; org-gdoc.el --- Sync between a single Org-mode file and a folder of docx files -*- lexical-binding: t; -*-

;; Copyright (C) 2024 Bala Ramadurai
;; Author: Bala Ramadurai <bala@balaramadurai.net>
;;         Assisted by Gemini 2.5 Pro
;; URL: https://github.com/balaramadurai/org-gdoc
;; Keywords: org, sync, docx, google-docs
;; Version: 0.1
;; Package-Requires: ((emacs "27.1"))

;; This file is not part of GNU Emacs.

;;; License:
;; Permission is hereby granted, free of charge, to any person obtaining a copy
;; of this software and associated documentation files (the "Software"), to deal
;; in the Software without restriction, including without limitation the rights
;; to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
;; copies of the Software, and to permit persons to whom the Software is
;; furnished to do so, subject to the following conditions:

;; The above copyright notice and this permission notice shall be included in all
;; copies or substantial portions of the Software.

;; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
;; IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
;; FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
;; AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
;; LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
;; OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
;; SOFTWARE.

;;; Commentary:
;; This package provides a bridge between a single, master Org-mode file and a
;; directory of .docx files, a common workflow for writers collaborating with
;; editors who use Google Docs or Microsoft Word. It uses a Python backend with
;; the `pandoc` utility to handle file conversions.

;;; Code:
(require 'org)
(require 'thingatpt)

(defgroup org-gdoc nil
  "Settings for the org-gdoc synchronization tool."
  :group 'org)

(defun org-gdoc--get-script-path ()
  "Robustly find the path to the `org_gdoc_sync.py` script.
This handles being run from source, from a straight.el `build` dir, or
when the user has customized `org-gdoc-python-script`."
  (if (and (boundp 'org-gdoc-python-script) org-gdoc-python-script)
      org-gdoc-python-script
    (let* ((el-file-path (or load-file-name buffer-file-name))
           (el-dir (file-name-directory el-file-path)))
      (cond
       ;; Standard source layout
       ((file-exists-p (expand-file-name "org_gdoc_sync.py" el-dir))
        (expand-file-name "org_gdoc_sync.py" el-dir))
       ;; straight.el layout: we are in build/, need to find repos/
       ((string-match-p "/straight/build/" el-file-path)
        (let* ((build-dir-name (file-name-nondirectory (directory-file-name el-dir)))
               (repos-dir (expand-file-name "../repos/" el-dir))
               (script-path (expand-file-name (concat build-dir-name "/org_gdoc_sync.py") repos-dir)))
          (if (file-exists-p script-path)
              script-path
            (error "Could not find python script in straight.el repos dir: %s" script-path))))
       (t (error "Could not automatically find org_gdoc_sync.py. Please set `org-gdoc-python-script`."))))))

(defcustom org-gdoc-python-script nil
  "The full path to the `org_gdoc_sync.py` script.
If nil, the script is assumed to be in the same directory as this .el file."
  :type '(file :must-match t)
  :group 'org-gdoc)

(defcustom org-gdoc-python-executable "python3"
  "The command to execute Python."
  :type 'string
  :group 'org-gdoc)

(defcustom org-gdoc-reference-doc nil
  "Path to a .docx file to use as a style reference for exports.
Create a .docx file, modify its styles (e.g., set the 'Normal'
style to use your preferred font and size), and point this
variable to it. This ensures consistent formatting in the
generated documents."
  :type '(file :must-match t)
  :group 'org-gdoc)

(defun org-gdoc--run-import (args message-template)
  "Internal helper to run the import process and handle errors."
  (let* ((sync-buffer-name "*org-gdoc-sync*")
         (output-buffer (get-buffer-create sync-buffer-name))
         (script-path (org-gdoc--get-script-path))
         (process-args (append (list script-path "import") args)))
    (with-current-buffer output-buffer (erase-buffer))
    (let ((exit-code (apply #'call-process org-gdoc-python-executable nil output-buffer nil process-args)))
      (if (= exit-code 0)
          (progn
            (message message-template)
            (when (get-buffer sync-buffer-name) (kill-buffer sync-buffer-name)))
        (pop-to-buffer sync-buffer-name)
        (error "Org-gdoc import failed. See *org-gdoc-sync* buffer for details.")))))

;;;###autoload
(defun org-gdoc-import-from-folder ()
  "Import all .docx files from a folder into the current Org buffer.
The function will only import documents that are not already
present in the file (based on the :SOURCE_FILE: property)."
  (interactive)
  (let ((folder (read-directory-name "Import from folder: "))
        (output-file (buffer-file-name)))
    (if (not output-file)
        (error "Buffer is not visiting a file")
      (message "Importing from %s..." folder)
      (org-gdoc--run-import (list "--folder" folder "--output" output-file)
                            (format "Successfully imported new documents from %s" folder)))))

;;;###autoload
(defun org-gdoc-insert-subtree-from-new-file ()
  "Import a single .docx file and insert its content as a new subtree at point."
  (interactive)
  (let ((source-file (read-file-name "Import .docx file: " nil nil t))
        (temp-file (make-temp-file "org-gdoc-import-")))
    (message "Importing %s..." source-file)
    (org-gdoc--run-import (list "--file" source-file "--output" temp-file)
                          "") ;; Don't show success message for this part
    (save-excursion
      (insert-file-contents temp-file)
      (delete-file temp-file))
    (message "Successfully inserted content from %s" source-file)))

;;;###autoload
(defun org-gdoc-export-subtree ()
  "Export the current subtree to its corresponding .docx file.
Uses the :SOURCE_FILE: property in the drawer to determine the
target file path."
  (interactive)
  (save-excursion
    (org-back-to-heading t)
    (let* ((source-file (org-entry-get (point) "SOURCE_FILE"))
           (start (point))
           (end (progn (org-end-of-subtree) (point)))
           (sync-buffer-name "*org-gdoc-sync*")
           (output-buffer (get-buffer-create sync-buffer-name))
           (script-path (org-gdoc--get-script-path)))
      (unless source-file
        (error "No :SOURCE_FILE: property found in the current subtree"))
      (message "Syncing subtree to %s..." source-file)
      (with-current-buffer output-buffer (erase-buffer))
      (let* ((args (list script-path "export" "--target-file" source-file "--quiet"))
             (full-args (if org-gdoc-reference-doc
                            (append args (list "--reference-doc" org-gdoc-reference-doc))
                          args))
             (exit-code (apply #'call-process-region start end org-gdoc-python-executable nil output-buffer nil full-args)))
        (if (= exit-code 0)
            (progn
              (message "Subtree successfully synced to %s" source-file)
              (when (get-buffer sync-buffer-name) (kill-buffer sync-buffer-name)))
          (pop-to-buffer sync-buffer-name)
          (error "Org-gdoc export failed. See *org-gdoc-sync* buffer for details."))))))

;;;###autoload
(defun org-gdoc-export-buffer ()
  "Export all top-level subtrees in the buffer to their .docx files."
  (interactive)
  (save-excursion
    (goto-char (point-min))
    (message "Exporting all subtrees in the buffer...")
    ;; `org-map-entries` is the robust way to iterate over headings.
    ;; We call `org-gdoc-export-subtree` for each level-1 heading.
    (org-map-entries #'org-gdoc-export-subtree "LEVEL=1" 'file)
    (message "Finished exporting all subtrees.")))

;;;###autoload
(defun org-gdoc-update-subtree-from-source ()
  "Update the current subtree by re-importing from its :SOURCE_FILE:."
  (interactive)
  (save-excursion
    (org-back-to-heading t)
    (let* ((source-file (org-entry-get (point) "SOURCE_FILE"))
           (temp-file (make-temp-file "org-gdoc-update-"))
           (start (point)))
      (unless source-file
        (error "No :SOURCE_FILE: property found in current subtree"))
      (message "Updating subtree from %s..." source-file)
      ;; 1. Import the docx content into a temporary org file.
      (org-gdoc--run-import (list "--file" source-file "--output" temp-file) "")
      ;; 2. Delete the current subtree.
      (org-end-of-subtree)
      (delete-region start (point))
      ;; 3. Insert the new content from the temp file.
      (goto-char start)
      (insert-file-contents temp-file)
      (delete-file temp-file)
      (message "Subtree successfully updated from %s." source-file))))

;;;###autoload
(defun org-gdoc-split-buffer-to-folder ()
  "Split the current Org buffer into multiple .docx files in a directory.
Each top-level heading in the buffer becomes a separate .docx file."
  (interactive)
  (when (buffer-modified-p)
    (when (y-or-n-p "Buffer is modified; save it first? ")
      (save-buffer)))
  (let* ((source-file (buffer-file-name))
         (output-folder (read-directory-name "Output folder for .docx files: ")))
    (if (not source-file)
        (error "Buffer is not visiting a file")
      (message "Splitting %s into %s..." (file-name-nondirectory source-file) output-folder)
      (let* ((script-path (org-gdoc--get-script-path))
             (args (list script-path
                           "split"
                           "--source-file" source-file
                           "--output-folder" output-folder))
             (sync-buffer-name "*org-gdoc-sync*")
             (output-buffer (get-buffer-create sync-buffer-name)))

        (when org-gdoc-reference-doc
          (setq args (append args (list "--reference-doc" org-gdoc-reference-doc))))

        (with-current-buffer output-buffer (erase-buffer))
        
        (let ((exit-code (apply #'call-process org-gdoc-python-executable nil output-buffer nil args)))
          (if (= exit-code 0)
              (progn
                (message "Successfully split buffer into .docx files in %s." output-folder)
                (when (get-buffer sync-buffer-name) (kill-buffer sync-buffer-name)))
            (pop-to-buffer sync-buffer-name)
            (error "Org-gdoc split failed. See *org-gdoc-sync* buffer for details.")))))))


(provide 'org-gdoc)

;;; org-gdoc.el ends here

