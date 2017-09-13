(in-package :clods-export)

(defvar *columns* nil
  "The set of columns defined for the current document.")

(defvar *last-column* nil
  "Pointer to the last cons cell of the columns list, to facilitate appending of new columns.")

(defclass column ()
  ((style :initform nil :initarg :style :reader col-style)
   (repeat :initform 1 :initarg :repeat :reader col-repeat)))

(defmacro with-body (() &body body)
  "This form encloses the actual content of the spreadsheet document.
The document consists of 1 to n tables that are written with the with-table macro."
  `(progn
     (unless (find *sheet-state* '(:start :fonts-defined :styles-defined))
       (error "with-body must be the last child of a with-spreadsheet form"))
     (with-element* ("office" "body")
       (with-element* ("office" "spreadsheet")
	 (let ((*sheet-state* :spreadsheet))
	   ,@body
	   (unless (eq *sheet-state* :table-seen)
	     (error "no data on spreadsheet - at least one with-table form is required")))
	 (setf *sheet-state* :end)))))

(defmacro with-table ((name) &body body)
  "Content of a single table (work sheet) on the document."
  (alexandria:with-gensyms (nsym)
    `(let ((,nsym ,name))
       (unless (find *sheet-state* '(:spreadsheet :table-seen))
	 (error "with-table must be a child of a with-body form"))
       ,(unless (typep name 'string)
		`(check-type ,nsym string))
       (with-element* ("table" "table")
	 (attribute* "table" "name" ,nsym)
	 (with-element* ("table" "title")
	   (text ,nsym))
	 (let ((*sheet-state* :table)
	       (*columns* nil)
	       (*last-column* nil))
	   ,@body)
	 (setf *sheet-state* :table-seen)))))

(defmacro with-header-columns (() &body body)
  "Header columns may be grouped together with this form.
They do not have a visual effect, but may carry a semantic meaning."
  `(with-element* ("table" "table-header-columns")
     (let ((*sheet-state* :columns-only))
       ,@body)))

(defun write-col-row-attrs (repeat style visibility cell-style)
  "Helper function for writing out attributes that are common to rows and columns."
  (when repeat
    (check-type repeat (integer 1 *))
    (attribute* "table" "number-columns-repeated" (princ-to-string repeat)))
  (when style
    (check-type style string)
    (attribute* "table" "style-name" style))
  (when visibility
    (check-type visibility (member :visible :collapse :filter))
    (attribute* "table" "visibility" (string-downcase visibility)))
  (when cell-style
    (check-type cell-style string)
    (attribute* "table" "default-cell-style-name" cell-style)))

(defun column (&key repeat style visibility cell-style)
  "Define a column (or, with the repeat argument, several similar columns)."
  (unless (find *sheet-state* '(:columns-only :table))
    (error "Columns must be defined inside a with-table form, and all columns must be defined before the first row."))
  (with-element* ("table" "table-column")
    (write-col-row-attrs repeat style visibility cell-style))
  (let ((col (make-instance 'column :style cell-style :repeat (or repeat 1))))
    (if *last-column*
	(setf (cdr *last-column*) (list col)
	      *last-column* (cdr *last-column*))
	(setf *columns* (list col)
	      *last-column* *columns*)))
  t)

(defmacro with-header-rows (() &body body)
  "Header rows may be grouped together with this form.
They do not have a visual effect, but may carry a semantic meaning."
  `(progn
     (unless (eq *sheet-state* :table)
       (error "Rows must be defined inside a with-table form, and header rows before others."))
     (with-element* ("table" "table-header-rows")
       (setf *sheet-state* :rows-only)
       ,@body)))

(defvar *current-column* nil
  "Track the current column when filling in data.")

(defvar *current-row-style* nil
  "Holds the default style specified for the current row.")

(defmacro with-row ((&key repeat style visibility cell-style) &body body)
  "Encloses a single row (or several similar, if the repeat arugment is used) on the table."
  `(progn
     (unless (find *sheet-state* '(:table :rows-only))
       (error "Rows must be defined inside a with-table form"))
     (with-element* ("table" "table-row")
       (write-col-row-attrs ,repeat ,style ,visibility ,cell-style)
       (multiple-value-prog1
	   (let ((*sheet-state* :cells-only)
		 (*current-column* (list *columns* 0))
		 (*current-row-style* ,cell-style))
	     ,@body)
	 (setf *sheet-state* :rows-only)))))

(defun cell (content &key style formula span-columns span-rows link)
  "Write out the contents of a data cell.

The cell may span several columns or rows, and it can contain a
formula or formatted data.  The content is formatted according to the
chosen data style (first cell style; if not specified, current row
style; if not specified, current column style; if not specified, an
error is signalled unless content is a string).

The following content types are supported:
* _real_ maps to ODS float, currency, or percentage type, according to the chosen style.
* _local-time:timestamp_ maps to ODS date or time type, according to the chosen style.
* the keywords :true and :false map to the ODS boolean type
* _string_ is the content as-is, without formatting or numeric value."
  (unless (eq *sheet-state* :cells-only)
    (error "All cells must be defined inside with-row forms."))
  (let* ((style-name (or style *current-row-style*
			 (and *current-column* (col-style (first (first *current-column*))))))
	 (style (gethash style-name *styles*))
	 (data-style-name (and style (slot-value style 'data-style)))
	 (data-style (or (gethash data-style-name *styles*) *default-data-style*)))
    ;; Advance the current column counter
    (dotimes (span (or span-columns 1))
      (when *current-column*
	(when (>= (incf (second *current-column*))
		  (col-repeat (first (first *current-column*))))
	  (setf (first *current-column*) (rest (first *current-column*))
		(second *current-column*) 0)
	  (unless (first *current-column*)
	    (setf *current-column* nil)))))
    ;; Verify that a valid style has been specified, and there is a
    ;; mapping to a data style
    (when (and style-name (not style))
      (error "Unknown style: ~s" style-name))
    (when (and data-style-name (eq data-style *default-data-style*))
      (error "Unknown data-style: ~s" data-style-name))
    (unless (or (null content) (stringp content))
      (unless style
	(error "Style is required for all non-string cells (content: ~s)" content))
      (unless data-style-name
	(error "data-style mapping is missing from style ~s" style-name)))
    (with-element* ("table" "table-cell")
      ;; Basic attributes
      (when (or style-name *current-row-style*)
	(check-type style-name (or null string))
	(attribute* "table" "style-name" (or style-name *current-row-style*)))
      (when formula
	(check-type formula string)
	(attribute* "table" "formula" formula))
      ;; Add span column/row info
      (when (or span-columns span-rows)
	(check-type span-columns (or null (integer 1 *)))
	(check-type span-rows (or null (integer 1 *)))
	(attribute* "table" "number-columns-spanned" (princ-to-string (or span-columns 1)))
	(attribute* "table" "number-rows-spanned" (princ-to-string (or span-rows 1))))
      ;; Format the cell according to the content
      (let ((text (etypecase content
		    (null nil)

		    (real
		     (check-type data-style (or number-number-style number-currency-style number-percentage-style))
		     (attribute* "office" "value-type" (ods-value-type data-style))
		     (attribute* "office" "value" (princ-number content))
		     (format-data data-style content))

		    (local-time:timestamp
		     (check-type data-style (or number-date-style number-time-style))
		     (if (typep data-style 'number-date-style)
			 (progn
			   (attribute* "office" "value-type" "date")
			   (attribute* "office" "date-value" (local-time:format-rfc3339-timestring nil content :use-zulu t :omit-time-part t)))
			 (progn
			   (attribute* "office" "value-type" "time")
			   (attribute* "office" "time-value" (remove-nsec
							     (local-time:format-rfc3339-timestring nil content :omit-date-part t :use-zulu t)))))
		     (format-data data-style content))

		    (keyword
		     (check-type content (member :true :false))
		     (attribute* "office" "value-type" "boolean")
		     (attribute* "office" "boolean-value" (string-downcase content))
		     (format-data data-style content))

		    (string
		     (attribute* "office" "value-type" "string")
		     (let ((fmt (and (typep data-style 'number-text-style)
				     (format-data data-style content))))
		       (unless (and fmt (string= fmt content))
			 (attribute* "office" "string-value" content))
		       fmt)))))

	(when text
	  (if (position #\newline text)
	      ;; multi-line
	      (loop for start = (or (and end (1+ end)) 0)
		    for end = (position #\newline text :start start)
		    for sub = (subseq text start end)
		    do (with-element* ("text" "p")
			 (if link
			     (with-element* ("text" "a")
			       (attribute* "xlink" "href" link)
			       (text sub))
			     (text sub)))
		    while end)
	      ;; single-line
	      (with-element* ("text" "p")
		(if link
		    (with-element* ("text" "a")
		      (attribute* "xlink" "href" link)
		      (text text))
		    (text text)))))))

    (when (and span-columns (> span-columns 1))
      (dotimes (i (1- span-columns))
	(with-element* ("table" "covered-table-cell"))))))

(defun covered-cell (&optional n)
  "Write out a set of covered cells.
When a cell spans several columns horizontally, the covered cells are
written automatically, but when spanning over rows vertically, the
user of the library is responsible of adding the covered cells where
the vertical spanning takes place."
  (dotimes (i n)
    (with-element* ("table" "covered-table-cell"))))

(defun cells (&rest content)
  "A convenience function for writing out a set of cells with no special formatting."
  (dolist (i content)
    (cell i)))
