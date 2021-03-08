(in-package :clods-export)

;;; The class hierarchy for holding the information of column/row/cell styles

(defclass base-style ()
  ((name :initarg :name :reader style-name)
   (data-style :initarg :data-style)
   (%-data-style :initarg :%-data-style)))

(defclass table-style (base-style)
  ())

(defclass column-style (base-style)
  ())

(defclass row-style (base-style)
  ())

(defclass cell-style (base-style)
  ())

(defmethod print-object ((obj base-style) stream)
  (print-unreadable-object (obj stream :type t :identity t)
    (prin1 (style-name obj) stream)))

(defmacro using-styles ((&key locale) &body body)
  "Specify the styles to be used in the document.  The default locale
can be specified; in this case, all styles get that locale
automatically."
  `(progn
     (unless (or (eq *sheet-state* :start)
		 (eq *sheet-state* :fonts-defined))
       (error "with-styles must appear as a child of a with-spreadsheet form, in the beginning or just after a with-fonts form"))
     (with-element* ("office" "automatic-styles")
       (let ((*sheet-state* :automatic-styles)
	     (*default-locale* (or ,locale *default-locale*)))
	 ,@body)
       (setf *sheet-state* :styles-defined))))

(defun ensure-style-state ()
  "Helper function to enforce state on the various style specs."
  (unless (eq *sheet-state* :automatic-styles)
    (error "styles can only be defined inside a with-styles form")))

(defun write-number-number (spec &key (normal-only t))
  "Helper function to write out a number specification.
Type of the number (normal/scientific/fraction) is deduced from the flags."
  (let* (not-normal not-frac not-sci
	 (attrs (loop for (kw data) on spec by #'cddr
		      collect (list kw
				    (ecase kw
				      (:min-integer-digits
				       (check-type data integer)
				       (princ-to-string data))
				      (:decimal-places
				       (check-type data integer)
				       (setf not-frac t)
				       (princ-to-string data))
				      (:decimal-replacement
				       (check-type data string)
				       (setf not-frac t not-sci t)
				       data)
				      (:display-factor
				       (check-type data real)
				       (setf not-frac t not-sci t)
				       (princ-number data))
				      (:number-grouping
				       (if data "true" "false"))
				      (:min-exponent-digits
				       (check-type data integer)
				       (setf not-normal t not-frac t)
				       (princ-to-string data))
				      ((:denominator-value :min-denominator-digits :min-numerator-digits)
				       (check-type data integer)
				       (setf not-normal t not-sci t)
				       (princ-to-string data)))))))
    (with-element* ("number" (cond ((not not-normal) "number")
				   (normal-only (error "invalid number specification: ~s" spec))
				   ((not not-sci) "scientific-number")
				   ((not not-frac) "fraction")
				   ((error "conflicting number specification: ~s" spec))))
      (dolist (i attrs)
	(attribute* "number" (string-downcase (first i))
		    (second i))))))

(defmacro defnumstyle (name (&rest args) (&rest init-list) &body body)
  "Helper macro for defining number styles."
  (let ((sym (intern (format nil "NUMBER-~A-STYLE" name))))
    `(defun ,sym (style-name ,@args text-properties (locale *default-locale*))
       ,(and (stringp (first body))
	     (first body))
       (ensure-style-state)
       (with-element* ("number" ,(concatenate 'string (string-downcase name) "-style"))
	 (attribute* "style" "name" style-name)
	 (when (locale-country locale)
	   (attribute* "number" "country" (locale-country locale)))
	 (write-text-properties text-properties)
	 ,@body)
       (when (gethash style-name *styles*)
	 (warn "Overriding style ~s with a new definition." style-name))
       (setf (gethash style-name *styles*)
	     (make-instance ',sym :name style-name ,@init-list :locale locale))
       style-name)))

(defnumstyle boolean (true false &key prefix suffix)
    (:true true :false false :prefix prefix :suffix suffix)
  "Specify a number:boolean-style.  The value is formatted as the
strings given in the true and false arguments."
  (when prefix
    (with-element* ("number" "text")
      (text prefix)))
  (with-element* ("number" "boolean"))
  (when suffix
    (with-element* ("number" "text")
      (text suffix))))

(defnumstyle currency (format &key)
    (:format format)
  "Specify a number:currencty-style.
The format is a plist of the tags :symbol, :text and :number; for
example, '(:number (:min-integer-digits 1 :decimal-places 2) :text \" \" :symbol \"€\")."
  (check-type format list)
  (unless (and (evenp (length format))
	       (<= (count :symbol format) 1)
	       (<= (count :text format) 1)
	       (= (count :number format) 1))
    (error "invalid currency style format ~s" format))
  (iter (for (kw data) on format by #'cddr)
	(ecase kw
	  (:symbol (with-element* ("number" "currency-symbol") (text data)))
	  (:text (with-element* ("number" "text") (text data)))
	  (:number (write-number-number data)))))

(defun date-or-time-style (date-p format)
  "Helper function for writing out a date or time style."
  (dolist (i format)
    (etypecase i
      (string (with-element* ("number" "text") (text i)))
      (keyword (let* ((sym (symbol-name i))
		      (dash (position #\- sym))
		      (style (and dash (string-downcase (subseq sym 0 dash))))
		      (field (and dash (string-downcase (subseq sym (1+ dash))))))
		 (unless (and dash
			      (find style '("long" "short" "am") :test #'string=)
			      (or (find field '("hours" "minutes" "seconds" "pm") :test #'string=)
				  (and date-p
				       (find field '("day" "month" "year" "era"
						     "day-of-week" "week-of-year" "quarter")
					     :test #'string=)))
			      (not (or (and (string= style "am")
					    (string/= style "pm"))
				       (and (string/= style "am")
					    (string= field "pm")))))
		   (error "invalid date/time field: ~s" i))
		 (with-element* ("number" field)
		   (attribute* "number" "style" style)))))))

(defnumstyle date (format &key)
    (:format format)
  "Specify a number:date-style.
The format is a list of the tags :[long/short]-[day/month/year/era/day-of-week/week-of-year/quarter/hours/minutes/seconds/am-pm]
and strings, for example '(:long-year \"-\" :long-month \"-\" :long-day)."
  (date-or-time-style t format))

(defnumstyle time (format &key)
    (:format format)
  "Specify a number:time-style.
The format is a list of the tags :[long/short]-[hours/minutes/seconds/am-pm]
and strings, for example '(:long-hours \":\" :long-minutes \":\" :long-seconds)."
  (date-or-time-style nil format))

(defnumstyle number (format &key prefix suffix)
    (:format format :prefix prefix :suffix suffix)
  "Specify a number:number-style. The format is a plist with the following tags:
* :min-integer-digits :number-grouping for all number styles,
* :decimal-replacement :display-factor for normal number styles,
* :decimal-places for normal and scientific styles,
* :min-exponent-digits for scientific styles,
* :denominator-value :min-denominator-digits :min-numerator-digits for fraction styles."
  (when prefix
    (with-element* ("number" "text")
      (text prefix)))
  (write-number-number format :normal-only nil)
  (when suffix
    (with-element* ("number" "text")
      (text suffix))))

(defnumstyle percentage (format &key prefix (suffix " %"))
    (:format format :prefix prefix :suffix suffix)
  "Specify a number:percentage-style. The format is the same as for number-number-style."
  (when prefix
    (with-element* ("number" "text")
      (text prefix)))
  (write-number-number format)
  (when suffix
    (with-element* ("number" "text")
      (text suffix))))

(defnumstyle text (&key prefix suffix)
    (:prefix prefix :suffix suffix)
  "Specify a number:text-style. Used for simple text content."
  (when prefix
    (with-element* ("number" "text")
      (text prefix)))
  (with-element* ("number" "text-content"))
  (when suffix
    (with-element* ("number" "text")
      (text suffix))))

(defmacro defstyle (name (&rest args) &body body)
  "Helper macro for style:style generators."
  `(defun ,(if (listp name) (first name) name)
       (name parent-style ,@args next-style data-style percentage-data-style)
     ,(and (stringp (first body))
	   (first body))
     (ensure-style-state)
     (with-element* ("style" "style")
       (attribute* "style" "name" name)
       (attribute* "style" "family" ,(if (listp name)
					(second name)
					(string-downcase (subseq (symbol-name name) 0 (position #\- (symbol-name name))))))
       (when parent-style
	 (attribute* "style" "parent-style-name" parent-style))
       (when next-style
	 (attribute* "style" "next-style-name" next-style))
       (when data-style
	 (attribute* "style" "data-style-name" data-style))
       (when percentage-data-style
	 (attribute* "style" "percentage-data-style-name" percentage-data-style))
       ,@body)
     (when (gethash name *styles*)
       (warn "Overriding style ~s with a new definition." name))
     (setf (gethash name *styles*)
	   (make-instance ',(if (listp name) (first name) name)
			  :name name :data-style data-style :%-data-style percentage-data-style))
     name))

(defun write-background (background)
  "Helper function for writing the background color specification."
  (when background
    (attribute* "fo" "background-color"
	  (if (eq background :transparent)
	      "transparent"
	      (if (valid-color background)
		  background
		  (error "invalid color specification: ~s" background))))))

(defstyle table-style (&key width rel-width align background)
  "Generator for <style:style style:family=\"table\">."
  (with-element* ("style" "table-properties")
    (when width
      (check-type width string)
      (attribute* "style" "width" width))
    (when rel-width
      (check-type rel-width string)
      (attribute* "style" "rel-width" rel-width))
    (when align
      (check-type align (member :left :center :right :margins))
      (attribute* "table" "align" (string-downcase align)))
    (write-background background)))

(defstyle (column-style "table-column") (&key width rel-width (use-optimal-width nil opt-width-supplied))
  "Generator for <style:style style:family=\"table-column\">."
  (with-element* ("style" "table-column-properties")
    (when width
      (check-type width string)
      (attribute* "style" "column-width" width))
    (when rel-width
      (check-type rel-width string)
      (attribute* "style" "rel-column-width" rel-width))
    (when opt-width-supplied
      (attribute* "style" "use-optimal-column-width" (if use-optimal-width "true" "false")))))

(defstyle (row-style "table-row") (&key height min-height (use-optimal-height nil opt-height-supplied) background)
  "Generator for <style:style style:family=\"table-row\">."
  (with-element* ("style" "table-row-properties")
    (when height
      (check-type height string)
      (attribute* "style" "row-height" height))
    (when min-height
      (check-type min-height string)
      (attribute* "style" "min-row-height" min-height))
    (when opt-height-supplied
      (attribute* "style" "use-optimal-row-height" (if use-optimal-height "true" "false")))
    (write-background background)))

(defstyle (cell-style "table-cell") (text-properties
				     &key horizontal-align vertical-align text-align-source background
				     border border-left border-top border-right border-bottom
				     (wrap nil wrap-supplied) rotation-angle)
  "Generator for <style:style style:family=\"table-cell\">."
  (with-element* ("style" "table-cell-properties")
    (when vertical-align
      (check-type vertical-align (member :top :middle :bottom :automatic))
      (attribute* "style" "vertical-align" (string-downcase vertical-align)))
    (when text-align-source
      (check-type text-align-source (member :fix :value-type))
      (attribute* "style" "text-align-source" (string-downcase text-align-source)))
    (write-background background)
    (flet ((write-border (dir which)
	     (let ((val (or which border)))
	       (when val
		 (check-type val list)
		 (unless (and (= (length val) 3)
			      (find (first val) *line-widths*)
			      (find (second val) *line-styles*)
			      (valid-color (third val)))
		   (error "invalid border definition: ~s" val))
		 (attribute* "fo" (format nil "border-~a" dir)
			     (format nil "~a ~a ~a"
				     (string-downcase (first val))
				     (string-downcase (second val))
				     (third val)))))))
      (write-border "left" border-left)
      (write-border "top" border-top)
      (write-border "right" border-right)
      (write-border "bottom" border-bottom))
    (when wrap-supplied
      (attribute* "fo" "wrap-option" (if wrap "wrap" "no-wrap")))
    (when rotation-angle
      (check-type rotation-angle (or string number))
      (attribute* "style" "rotation-angle" (princ-to-string rotation-angle))))
  (when text-properties
    (write-text-properties text-properties))
  (when horizontal-align
    (check-type horizontal-align (member :start :center :end :justify :left :right))
    (with-element* ("style" "paragraph-properties")
      (attribute* "fo" "text-align" (string-downcase horizontal-align)))))
