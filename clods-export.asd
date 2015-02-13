(in-package #:cl)
(defpackage #:clods-export-system (:use #:asdf #:cl))
(in-package #:clods-export-system)

(asdf:defsystem clods-export
  :name "clods-export"
  :author "Jussi Lahdenniemi <jussi@lahdenniemi.fi>"
  :maintainer "Jussi Lahdenniemi <jussi@lahdenniemi.fi>"
  :license "MIT"
  :description "Common Lisp OpenDocument spreadsheet export library"
  :encoding :utf-8
  :depends-on (alexandria iterate xmlw zip cl-fad)
  :components
  ((:module clods
    :pathname ""
    :components ((:file "package")
		 (:file "ods" :depends-on ("package"))
		 (:file "fix-zip")))))
