(in-package :clods-export)

;;; Namespaces used in ODS documents

(defparameter *ns-table* "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
(defparameter *ns-office* "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
(defparameter *ns-text* "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
(defparameter *ns-style* "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
(defparameter *ns-fo* "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
(defparameter *ns-of* "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
(defparameter *ns-xlink* "http://www.w3.org/1999/xlink")
(defparameter *ns-number* "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
(defparameter *ns-svg* "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
(defparameter *ns-meta* "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
(defparameter *ns-dc* "http://purl.org/dc/elements/1.1/")

;;; Convenience macro for referencing all the namespaces

(defmacro with-ods-namespaces (() &body body)
  `(with-namespace ("table" *ns-table*)
     (with-namespace ("office" *ns-office*)
       (with-namespace ("text" *ns-text*)
	 (with-namespace ("style" *ns-style*)
	   (with-namespace ("fo" *ns-fo*)
	     (with-namespace ("of" *ns-of*)
	       (with-namespace ("xlink" *ns-xlink*)
		 (with-namespace ("number" *ns-number*)
		   (with-namespace ("svg" *ns-svg*)
		     ,@body))))))))))
