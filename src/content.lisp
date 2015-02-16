(in-package :clods-export)

(defmacro with-body (() &body body)
  "This form encloses the actual content of the spreadsheet document.
The document consists of 1 to n tables that are written with the with-table macro."
  `(progn
     (unless (find *sheet-state* '(:start :fonts-defined :styles-defined))
       (error "with-body must be the last child of a with-spreadsheet form"))
     (with-tag ((*ns-office* "body"))
       (with-tag ((*ns-office* "spreadsheet"))
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
       (with-tag ((*ns-table* "table"))
	 (attr (*ns-table* "name") ,nsym)
	 (tag (*ns-table* "title") ,nsym)
	 (let ((*sheet-state* :table))
	   ,@body)
	 (setf *sheet-state* :table-seen)))))

(defmacro with-header-columns (() &body body)
  "Header columns may be grouped together with this form.
They do not have a visual effect, but may carry a semantic meaning."
  `(with-tag ((*ns-table* "table-header-columns"))
     (let ((*sheet-state* :columns-only))
       ,@body)))

(defun write-col-row-attrs (repeat style visibility cell-style)
  "Helper function for writing out attributes that are common to rows and columns."
  (when repeat
    (check-type repeat (integer 1 *))
    (attr (*ns-table* "number-columns-repeated") repeat))
  (when style
    (check-type style string)
    (attr (*ns-table* "style-name") style))
  (when visibility
    (check-type visibility (member :visible :collapse :filter))
    (attr (*ns-table* "visibility") (string-downcase visibility)))
  (when cell-style
    (check-type cell-style string)
    (attr (*ns-table* "default-cell-style-name") cell-style)))

(defun column (&key repeat style visibility cell-style)
  "Define a column (or, with the repeat argument, several similar columns)."
  (unless (find *sheet-state* '(:columns-only :table))
    (error "Columns must be defined inside a with-table form, and all columns must be defined before the first row."))
  (with-tag ((*ns-table* "table-column"))
    (write-col-row-attrs repeat style visibility cell-style)))

(defmacro with-header-rows (() &body body)
  "Header rows may be grouped together with this form.
They do not have a visual effect, but may carry a semantic meaning."
  `(progn
     (unless (eq *sheet-state* :table)
       (error "Rows must be defined inside a with-table form, and header rows before others."))
     (with-tag ((*ns-table* "table-header-rows"))
       (setf *sheet-state* :rows-only)
       ,@body)))

(defmacro with-row ((&key repeat style visibility cell-style) &body body)
  "Encloses a single row (or several similar, if the repeat arugment is used) on the table."
  `(progn
     (unless (find *sheet-state* '(:table :rows-only))
       (error "Rows must be defined inside a with-table form"))
     (with-tag ((*ns-table* "table-row"))
       (write-col-row-attrs ,repeat ,style ,visibility ,cell-style)
       (let ((*sheet-state* :cells-only))
	 ,@body))))

(defun cell (content &key style formula span-columns span-rows link value-type raw-value)
  "Write out the contents of a data cell.
The cell may span several columns or rows, and it can contain a formula or formatted data.
ODS supports a fixed set of data types that are reflected in the choice for the value-type argument:
* :float - all numbers, also integers, are floats; raw-value should be a real.
* :percentage - much like :float; semantic difference only.
* :currency - as above, semantic difference to :float only.
* :date - raw-value should be an ISO8601 formatted date.
* :time - raw-value should be an ISO8601 formatted time.
* :boolean - raw-value should be \"true\" or \"false\".
* :string - raw-value not needed, content is the string."
  (unless (eq *sheet-state* :cells-only)
    (error "All cells must be defined inside with-row forms."))
  (when content
    (check-type content string))
  (with-tag ((*ns-table* "table-cell"))
    (when style
      (check-type style string)
      (attr (*ns-table* "style-name") style))
    (when formula
      (check-type formula string)
      (attr (*ns-table* "formula") formula))
    (when (or content raw-value)
      (when value-type
	(check-type value-type (member :float :percentage :currency :date :time :boolean :string)))
      (attr (*ns-office* "value-type") (if value-type
					   (string-downcase value-type)
					   (typecase (or raw-value content)
					     (real "float")
					     (t "string"))))
      (unless (or (null value-type)
		  (eq value-type :string))
	(attr (*ns-office* (case (or value-type :float)
			     ((:float :percentage :currency) "value")
			     (:date "date-value")
			     (:time "time-value")
			     (:boolean "boolean-value")))
	      (cond ((realp raw-value)
		     (princ-number raw-value))
		    (raw-value
		     (princ-to-string raw-value))
		    (content)))))
    (when (or span-columns span-rows)
      (check-type span-columns (or null (integer 1 *)))
      (check-type span-rows (or null (integer 1 *)))
      (attr (*ns-table* "number-columns-spanned") (princ-to-string (or span-columns 1)))
      (attr (*ns-table* "number-rows-spanned") (princ-to-string (or span-rows 1))))
    (when content
      (with-tag ((*ns-text* "p") :compact t)
	(if link
	    (with-tag ((*ns-text* "a"))
	      (attr (*ns-xlink* "href") link)
	      (content content))
	    (content content)))))
  (when (and span-columns (> span-columns 1))
    (dotimes (i (1- span-columns))
      (tag (*ns-table* "covered-table-cell"))))
  (when (and span-rows (> span-rows 1))
    (error "span-rows not implemented properly yet")))

(defun cells (content &rest options)
  "A convenience function for writing out a set of similarly formatted cells."
  (dolist (i content)
    (apply #'cell i options)))
