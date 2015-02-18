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
     (with-tag ((*ns-office* "automatic-styles"))
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
    (with-tag ((*ns-number* (cond ((not not-normal) "number")
				  (normal-only (error "invalid number specification: ~s" spec))
				  ((not not-sci) "scientific-number")
				  ((not not-frac) "fraction")
				  ((error "conflicting number specification: ~s" spec)))))
      (dolist (i attrs)
	(attr (*ns-number* (string-downcase (first i)))
	      (second i))))))

(defmacro defnumstyle (name (&rest args) (&rest init-list) &body body)
  "Helper macro for defining number styles."
  (let ((sym (intern (format nil "NUMBER-~A-STYLE" name))))
    `(defun ,sym (style-name ,@args text-properties (locale *default-locale*))
       ,(and (stringp (first body))
	     (first body))
       (ensure-style-state)
       (with-tag ((*ns-number* ,(concatenate 'string (string-downcase name) "-style")))
	 (attr (*ns-style* "name") style-name)
	 (when (locale-country locale)
	   (attr (*ns-number* "country") (locale-country locale)))
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
    (tag (*ns-number* "text") prefix))
  (tag (*ns-number* "boolean"))
  (when suffix
    (tag (*ns-number* "text") suffix)))

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
	  (:symbol (tag (*ns-number* "currency-symbol") data))
	  (:text (tag (*ns-number* "text") data))
	  (:number (write-number-number data)))))

(defun date-or-time-style (date-p format)
  "Helper function for writing out a date or time style."
  (dolist (i format)
    (etypecase i
      (string (tag (*ns-number* "text") i))
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
		 (with-tag ((*ns-number* field)) (attr (*ns-number* "style") style)))))))

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
    (tag (*ns-number* "text") prefix))
  (write-number-number format :normal-only nil)
  (when suffix
    (tag (*ns-number* "text") suffix)))

(defnumstyle percentage (format &key)
    (:format format)
  "Specify a number:percentage-style. The format is the same as for number-number-style."
  (write-number-number format)
  (tag (*ns-number* "text") " %"))

(defnumstyle text (&key prefix suffix)
    (:prefix prefix :suffix suffix)
  "Specify a number:text-style. Used for simple text content."
  (when prefix
    (tag (*ns-number* "text") prefix))
  (tag (*ns-number* "text-content"))
  (when suffix
    (tag (*ns-number* "text") suffix)))

(defmacro defstyle (name (&rest args) &body body)
  "Helper macro for style:style generators."
  `(defun ,(if (listp name) (first name) name)
       (name parent-style ,@args next-style data-style percentage-data-style)
     ,(and (stringp (first body))
	   (first body))
     (ensure-style-state)
     (with-tag ((*ns-style* "style"))
       (attr (*ns-style* "name") name)
       (attr (*ns-style* "family") ,(if (listp name)
					(second name)
					(string-downcase (subseq (symbol-name name) 0 (position #\- (symbol-name name))))))
       (when parent-style
	 (attr (*ns-style* "parent-style-name") parent-style))
       (when next-style
	 (attr (*ns-style* "next-style-name") next-style))
       (when data-style
	 (attr (*ns-style* "data-style-name") data-style))
       (when percentage-data-style
	 (attr (*ns-style* "percentage-data-style-name") percentage-data-style))
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
    (attr (*ns-fo* "background-color")
	  (if (eq background :transparent)
	      "transparent"
	      (if (valid-color background)
		  background
		  (error "invalid color specification: ~s" background))))))

(defstyle table-style (&key width rel-width align background)
  "Generator for <style:style style:family=\"table\">."
  (with-tag ((*ns-style* "table-properties"))
    (when width
      (check-type width string)
      (attr (*ns-style* "width") width))
    (when rel-width
      (check-type rel-width string)
      (attr (*ns-style* "rel-width") rel-width))
    (when align
      (check-type align (member :left :center :right :margins))
      (attr (*ns-table* "align") (string-downcase align)))
    (write-background background)))

(defstyle (column-style "table-column") (&key width rel-width (use-optimal-width nil opt-width-supplied))
  "Generator for <style:style style:family=\"table-column\">."
  (with-tag ((*ns-style* "table-column-properties"))
    (when width
      (check-type width string)
      (attr (*ns-style* "column-width") width))
    (when rel-width
      (check-type rel-width string)
      (attr (*ns-style* "rel-column-width") rel-width))
    (when opt-width-supplied
      (attr (*ns-style* "use-optimal-column-width") (if use-optimal-width "true" "false")))))

(defstyle (row-style "table-row") (&key height min-height (use-optimal-height nil opt-height-supplied) background)
  "Generator for <style:style style:family=\"table-row\">."
  (with-tag ((*ns-style* "table-row-properties"))
    (when height
      (check-type height string)
      (attr (*ns-style* "row-height") height))
    (when min-height
      (check-type min-height string)
      (attr (*ns-style* "min-row-height") min-height))
    (when opt-height-supplied
      (attr (*ns-style* "use-optimal-row-height") (if use-optimal-height "true" "false")))
    (write-background background)))

(defstyle (cell-style "table-cell") (text-properties
				     &key horizontal-align vertical-align text-align-source background
				     border border-left border-top border-right border-bottom
				     (wrap nil wrap-supplied))
  "Generator for <style:style style:family=\"table-cell\">."
  (with-tag ((*ns-style* "table-cell-properties"))
    (when vertical-align
      (check-type vertical-align (member :top :middle :bottom :automatic))
      (attr (*ns-style* "vertical-align") (string-downcase vertical-align)))
    (when text-align-source
      (check-type text-align-source (member :fix :value-type))
      (attr (*ns-style* "text-align-source") (string-downcase text-align-source)))
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
		 (attr (*ns-fo* (format nil "border-~a" dir))
		       (format nil "~a ~a ~a"
			       (string-downcase (first val))
			       (string-downcase (second val))
			       (third val)))))))
      (write-border "left" border-left)
      (write-border "top" border-top)
      (write-border "right" border-right)
      (write-border "bottom" border-bottom))
    (when wrap-supplied
      (attr (*ns-fo* "wrap-option") (if wrap "wrap" "no-wrap"))))
  (when text-properties
    (write-text-properties text-properties))
  (when horizontal-align
    (check-type horizontal-align (member :start :center :end :justify :left :right))
    (with-tag ((*ns-style* "paragraph-properties"))
      (attr (*ns-fo* "text-align") (string-downcase horizontal-align)))))
