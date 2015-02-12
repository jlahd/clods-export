(in-package :clods)

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

(defun make-ods-styles (str)
  (writing-xml (str :standalone t)
    (with-tag ((*ns-office* "document-styles"))
      (add-ods-namespaces)
      (attr (*ns-office* "version") "1.2")
      (with-tag ((*ns-office* "font-face-decls")))
      (with-tag ((*ns-office* "styles")))
      (with-tag ((*ns-office* "automatic-styles")))
      (with-tag ((*ns-office* "master-styles"))))))

(defun make-ods-content (str)
  (writing-xml (str :standalone t)
    (with-tag ((*ns-office* "document-content"))
      (add-ods-namespaces)
      (attr (*ns-office* "version") "1.2")
      (with-tag ((*ns-office* "font-face-decls")))
      (with-tag ((*ns-office* "automatic-styles")))
      (with-tag ((*ns-office* "body"))
	(with-tag ((*ns-office* "spreadsheet"))
	  (with-tag ((*ns-table* "table"))
	    (attr (*ns-table* "name") "Vuokrasopimukset")
	    (with-tag ((*ns-table* "table-column"))
	      (attr (*ns-table* "number-colum*ns-repeated*") "4"))
	    (with-tag ((*ns-table* "table-row"))
	      (with-tag ((*ns-table* "table-cell"))
		(attr (*ns-office* "value-type") "string")
		(tag (*ns-text* "p") "Vuokanantaja"))
	      (with-tag ((*ns-table* "table-cell"))
		(attr (*ns-office* "value-type") "string")
		(tag (*ns-text* "p") "Tilan numero"))
	      (with-tag ((*ns-table* "table-cell"))
		(attr (*ns-office* "value-type") "string")
		(tag (*ns-text* "p") "Pinta-ala"))
	      (with-tag ((*ns-table* "table-cell"))
		(attr (*ns-office* "value-type") "string")
		(tag (*ns-text* "p") "Päiväys")))))))))

(defun make-ods-manifest (str)
  (let ((mfns (namespace "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")))
    (writing-xml (str :standalone t)
      (with-tag ((mfns "manifest"))
	(attr (*xmlns* "manifest") mfns)
	(with-tag ((mfns "file-entry"))
	  (attr (mfns "full-path") "/")
	  (attr (mfns "media-type") "application/vnd.oasis.opendocument.spreadsheet"))
	(with-tag ((mfns "file-entry"))
	  (attr (mfns "full-path") "styles.xml")
	  (attr (mfns "media-type") "text/xml"))
	(with-tag ((mfns "file-entry"))
	  (attr (mfns "full-path") "content.xml")
	  (attr (mfns "media-type") "text/xml"))
	(with-tag ((mfns "file-entry"))
	  (attr (mfns "full-path") "meta.xml")
	  (attr (mfns "media-type") "text/xml"))))))

(defun make-ods-meta (str)
  (let ((date (multiple-value-bind (sec min hour day month year)
		  (decode-universal-time (get-universal-time) 0)
		(format nil "~4,'0d-~2,'0d-~2,'0dT~2,'0d:~2,'0d:~2,'0dZ"
			year month day hour min sec))))
    (writing-xml (str :standalone t)
      (with-tag ((*ns-office* "document-meta"))
	(attr (*xmlns* "office") *ns-office*)
	(attr (*xmlns* "meta") *ns-meta*)
	(attr (*xmlns* "dc") *ns-dc*)
	(attr (*xmlns* "xlink") *ns-xlink*)
	(attr (*ns-office* "version") "1.2")
	(with-tag ((*ns-office* "meta"))
	  (tag (*ns-meta* "generator") "Ahma")
	  (tag (*ns-meta* "initial-creator") "Jussi Lahdenniemi")
	  (tag (*ns-dc* "creator") "Jussi Lahdenniemi")
	  (tag (*ns-meta* "creation-date") date)
	  (tag (*ns-dc* "date") date))))))

(defun export-to-ods (file)
  (let ((mimetype (make-temporary-filename))
	(styles (make-temporary-filename))
	(content (make-temporary-filename))
	(manifest (make-temporary-filename))
	(meta (make-temporary-filename)))
    (with-open-file (f mimetype :direction :output :if-exists :supersede :if-does-not-exist :create)
      (write-string "application/vnd.oasis.opendocument.spreadsheet" f))
    (with-open-file (f styles :direction :output :if-exists :supersede :if-does-not-exist :create)
      (make-ods-styles f))
    (with-open-file (f content :direction :output :if-exists :supersede :if-does-not-exist :create)
      (make-ods-content f))
    (with-open-file (f manifest :direction :output :if-exists :supersede :if-does-not-exist :create)
      (make-ods-manifest f))
    (with-open-file (f meta :direction :output :if-exists :supersede :if-does-not-exist :create)
      (make-ods-meta f))
    (zip:with-output-to-zipfile (zip file :if-exists :supersede)
      (flet ((add-file (name path)
	       (with-open-file (f path :direction :input :element-type '(unsigned-byte 8))
		 (zip:write-zipentry zip name f))))
	(add-file "mimetype" mimetype)
	(add-file "styles.xml" styles)
	(add-file "content.xml" content)
	(add-file "META-INF/manifest.xml" manifest)
	(add-file "meta.xml" meta)))))
