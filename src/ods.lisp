(in-package :clods-export)

(defun make-ods-styles ()
  "Create an empty styles.xml document"
  (cl-fad:with-output-to-temporary-file (str)
    (with-xml-output (make-character-stream-sink str)
      (with-namespace ("office" *ns-office*)
	(with-element* ("office" "document-styles")
	  (attribute* "office" "version" "1.2"))))))

(defun make-ods-manifest ()
  "Create a manifest document for files in the ODS zip"
  (cl-fad:with-output-to-temporary-file (str)
    (with-xml-output (make-character-stream-sink str)
      (with-namespace ("mfns" "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
	(with-element* ("mfns" "manifest")
	  (with-element* ("mfns" "file-entry")
	    (attribute* "mfns" "full-path" "/")
	    (attribute* "mfns" "media-type" "application/vnd.oasis.opendocument.spreadsheet"))
	  (with-element* ("mfns" "file-entry")
	    (attribute* "mfns" "full-path" "styles.xml")
	    (attribute* "mfns" "media-type" "text/xml"))
	  (with-element* ("mfns" "file-entry")
	    (attribute* "mfns" "full-path" "content.xml")
	    (attribute* "mfns" "media-type" "text/xml"))
	  (with-element* ("mfns" "file-entry")
	    (attribute* "mfns" "full-path" "meta.xml")
	    (attribute* "mfns" "media-type" "text/xml")))))))

(defun make-ods-meta (generator creator)
  "Create a meta.xml document, specifying a generator and creator"
  (cl-fad:with-output-to-temporary-file (str)
    (let ((date (remove-nsec (local-time:format-rfc3339-timestring nil (local-time:now) :use-zulu t))))
      (with-xml-output (make-character-stream-sink str)
	(with-namespace ("office" *ns-office*)
	  (with-namespace ("meta" *ns-meta*)
	    (with-namespace ("dc" *ns-dc*)
	      (with-element* ("office" "document-meta")
		(attribute* "office" "version" "1.2")
		(with-element* ("office" "meta")
		  (with-element* ("meta" "generator")
		    (text generator))
		  (with-element* ("meta" "initial-creator")
		    (text creator))
		  (with-element* ("dc" "creator")
		    (text creator))
		  (with-element* ("meta" "creation-date")
		    (text date))
		  (with-element* ("dc" "date")
		    (text date)))))))))))

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
			    (with-xml-output (make-character-stream-sink ,str)
			      (with-ods-namespaces ()
				(with-element* ("office" "document-content")
				  (attribute* "office" "version" "1.2")
				  (let ((*sheet-state* :start)
					(*styles* (make-hash-table :test 'equal)))
				    ,@body
				    (unless (eq *sheet-state* :end)
				      (error "no data on spreadsheet - a with-body form is required"))))))))
		    ,zip-files)
	      ;; The other files are written as simply as possible.
	      (push (cons "styles.xml" (make-ods-styles)) ,zip-files)
	      (push (cons "meta.xml" (make-ods-meta ,generator ,creator)) ,zip-files)
	      (push (cons "META-INF/manifest.xml" (make-ods-manifest)) ,zip-files)
	      (push (cons "mimetype" (make-ods-mimetype)) ,zip-files)
	      (zip-ods ,file ,zip-files))
	 (dolist (,i ,zip-files)
	   (delete-file (cdr ,i)))))))
