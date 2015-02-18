(in-package :clods-export)

(defun make-ods-styles ()
  "Create an empty styles.xml document"
  (cl-fad:with-output-to-temporary-file (str)
    (writing-xml (str :standalone t)
      (with-tag ((*ns-office* "document-styles"))
	(add-ods-namespaces)
	(attr (*ns-office* "version") "1.2")))))

(defun make-ods-manifest ()
  "Create a manifest document for files in the ODS zip"
  (cl-fad:with-output-to-temporary-file (str)
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
	    (attr (mfns "media-type") "text/xml")))))))

(defun make-ods-meta (generator creator)
  "Create a meta.xml document, specifying a generator and creator"
  (cl-fad:with-output-to-temporary-file (str)
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
	    (tag (*ns-meta* "generator") generator)
	    (tag (*ns-meta* "initial-creator") creator)
	    (tag (*ns-dc* "creator") creator)
	    (tag (*ns-meta* "creation-date") date)
	    (tag (*ns-dc* "date") date)))))))

(defun make-ods-mimetype ()
  "Create a mimetype file for the ODS zip"
  (cl-fad:with-output-to-temporary-file (str)
    (write-string "application/vnd.oasis.opendocument.spreadsheet" str)))

(defun zip-ods (out files)
  "Package a set of files into a zip archive"
  (zip:with-output-to-zipfile (zip out :if-exists :supersede)
    (dolist (file files)
      (with-open-file (f (cdr file) :direction :input :element-type '(unsigned-byte 8))
	(zip:write-zipentry zip (car file) f)))))

(defvar *sheet-state* nil
  "Tracks the state of the document being constructed.
Various forms appearing in a with-spreadsheet form must be specified
in the correct order; if the order is violated, an error is signalled.")

(defvar *styles* nil
  "A dictionary of all defined styles in the document.")

(defmacro with-spreadsheet ((file &key (generator "clods") (creator "clods")) &body body)
  "The main interface macro for CLODS-EXPORT.
Creates a ODS document out of the body of the form and stores it into
the specified file. The optional keyword arguments generator and
creator are written in the document's metadata."
  (alexandria:with-gensyms (zip-files i str)
    `(let (,zip-files)
       (unwind-protect
	    (progn
	      ;; content.xml is the only variable part.
	      (push (cons "content.xml"
			  (cl-fad:with-output-to-temporary-file (,str)
			    (writing-xml (,str :standalone t)
			      (with-tag ((*ns-office* "document-content"))
				(add-ods-namespaces)
				(attr (*ns-office* "version") "1.2")
				(let ((*sheet-state* :start)
				      (*styles* (make-hash-table :test 'equal)))
				  ,@body
				  (unless (eq *sheet-state* :end)
				    (error "no data on spreadsheet - a with-body form is required")))))))
		    ,zip-files)
	      ;; The other files are written as simply as possible.
	      (push (cons "styles.xml" (make-ods-styles)) ,zip-files)
	      (push (cons "meta.xml" (make-ods-meta ,generator ,creator)) ,zip-files)
	      (push (cons "META-INF/manifest.xml" (make-ods-manifest)) ,zip-files)
	      (push (cons "mimetype" (make-ods-mimetype)) ,zip-files)
	      (zip-ods ,file ,zip-files))
	 (dolist (,i ,zip-files)
	   (delete-file (cdr ,i)))))))
