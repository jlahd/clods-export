(in-package :clods-export)

;;; Namespaces used in ODS documents

(defparameter *ns-table* (namespace "urn:oasis:names:tc:opendocument:xmlns:table:1.0"))
(defparameter *ns-office* (namespace "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
(defparameter *ns-text* (namespace "urn:oasis:names:tc:opendocument:xmlns:text:1.0"))
(defparameter *ns-style* (namespace "urn:oasis:names:tc:opendocument:xmlns:style:1.0"))
(defparameter *ns-fo* (namespace "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"))
(defparameter *ns-of* (namespace "urn:oasis:names:tc:opendocument:xmlns:of:1.2"))
(defparameter *ns-xlink* (namespace "http://www.w3.org/1999/xlink"))
(defparameter *ns-number* (namespace "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"))
(defparameter *ns-svg* (namespace "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"))
(defparameter *ns-meta* (namespace "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"))
(defparameter *ns-dc* (namespace "http://purl.org/dc/elements/1.1/"))

;;; Convenience function for referencing all the namespaces

(defun add-ods-namespaces ()
  (attr (*xmlns* "table") *ns-table*)
  (attr (*xmlns* "office") *ns-office*)
  (attr (*xmlns* "text") *ns-text*)
  (attr (*xmlns* "style") *ns-style*)
  (attr (*xmlns* "fo") *ns-fo*)
  (attr (*xmlns* "of") *ns-of*)
  (attr (*xmlns* "xlink") *ns-xlink*)
  (attr (*xmlns* "number") *ns-number*)
  (attr (*xmlns* "svg") *ns-svg*))
