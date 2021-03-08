(in-package #:cl)
(defpackage #:clods-export-system (:use #:asdf #:cl))
(in-package #:clods-export-system)

(asdf:defsystem clods-export
  :name "clods-export"
  :version "1.0.2"
  :author "Jussi Lahdenniemi <jussi@lahdenniemi.fi>"
  :maintainer "Jussi Lahdenniemi <jussi@lahdenniemi.fi>"
  :license "MIT"
  :description "Common Lisp OpenDocument spreadsheet export library"
  :encoding :utf-8
  :depends-on (alexandria iterate local-time cxml zip cl-fad)
  :components
  ((:module clods
    :pathname "src"
    :components ((:file "package")
		 (:file "namespaces")
		 (:file "util")
		 (:file "locale")
		 (:file "ods" :depends-on ("package" "namespaces" "util"))
		 (:file "fonts" :depends-on ("ods"))
		 (:file "text" :depends-on ("ods" "fonts"))
		 (:file "numbers" :depends-on ("ods" "locale"))
		 (:file "styles" :depends-on ("ods" "text"))
		 (:file "content" :depends-on ("ods"))))))
