(in-package :clods-export)

;;; The class hierarchy for holding the information of number styles

(defclass number-base-style ()
  ((name :initarg :name :reader style-name)
   (locale :initarg :locale)))

(defclass number-boolean-style (number-base-style)
  ((prefix :initarg :prefix)
   (suffix :initarg :suffix)
   (true :initarg :true)
   (false :initarg :false)))

(defclass number-currency-style (number-base-style)
  ((format :initarg :format)))

(defclass number-date-time-style (number-base-style)
  ((format :initarg :format)))

(defclass number-date-style (number-date-time-style)
  ())

(defclass number-time-style (number-date-time-style)
  ())

(defclass number-number-style (number-base-style)
  ((format :initarg :format)
   (prefix :initarg :prefix)
   (suffix :initarg :suffix)))

(defclass number-percentage-style (number-number-style)
  ())

(defclass number-text-style (number-base-style)
  ((prefix :initarg :prefix)
   (suffix :initarg :suffix)))

(defmethod print-object ((obj number-base-style) stream)
  (print-unreadable-object (obj stream :type t :identity t)
    (prin1 (style-name obj) stream)))

(defparameter *default-data-style* (make-instance 'number-text-style :prefix nil :suffix nil))

(defgeneric ods-value-type (style)
  (:documentation "Retrieve the ODS value-type for the number-*-style."))

(defgeneric format-data (style data)
  (:documentation "Format data according to a number-*-style."))

(defmethod ods-value-type ((style number-boolean-style))
  "boolean")

(defmethod ods-value-type ((style number-currency-style))
  "currency")

(defmethod ods-value-type ((style number-date-style))
  "date")

(defmethod ods-value-type ((style number-time-style))
  "time")

(defmethod ods-value-type ((style number-number-style))
  "float")

(defmethod ods-value-type ((style number-percentage-style))
  "percentage")

(defmethod ods-value-type ((style number-text-style))
  "string")

(defmethod format-data ((style number-boolean-style) data)
  (with-slots (true false) style
    (ecase data
      (:true true)
      (:false false))))

(defun format-number (num fmt locale)
  "Format a number according to number formatting flags and the given locale."
  (let #.(mapcar #'(lambda (kw)
		     `(,(intern (symbol-name kw) :clods-export)
			(second (member ,kw fmt))))
		 '(:min-integer-digits
		   :decimal-places
		   :decimal-replacement
		   :display-factor
		   :number-grouping
		   :min-exponent-digits
		   :denominator-value
		   :min-denominator-digits
		   :min-numerator-digits))

       (cond (min-exponent-digits
	      ;; scientific
	      (let* ((stem (if decimal-places
			       (format nil "~,v,v,,,,'Ee" decimal-places min-exponent-digits (abs num))
			       (format nil "~,,v,,,,'Ee" min-exponent-digits (abs num))))
		     (dot (position #\. stem)))
		(setf (char stem dot) (decimal-separator locale))
		(concatenate 'string
			     (if (< num 0) "-" "")
			     (if (and min-integer-digits
				      (< dot min-integer-digits))
				 (make-string (- min-integer-digits dot) :initial-element #\0)
				 "")
			     stem)))
	     ((or denominator-value min-denominator-digits min-numerator-digits)
	      ;; fraction
	      (let* ((frac (rationalize num))
		     (den (or denominator-value (denominator frac)))
		     (num (if denominator-value
			      (round (* num denominator-value))
			      (numerator frac))))
		(when min-denominator-digits
		  (iter (for i below (- min-denominator-digits (length (princ-to-string den))))
			(setf den (* den 10)
			      num (* num 10))))
		(if (and min-integer-digits
			 (> min-integer-digits 0)
			 (>= (abs num) den))
		    (multiple-value-bind (wh fr)
			(truncate num den)
		      (format nil "~vd ~v,'0d/~d"
			      min-integer-digits wh
			      (or min-numerator-digits 0) (mod fr den) den))
		    (format nil "~v,'0d/~d" (or min-numerator-digits 0) num den))))
	     (t
	      ;; normal
	      (let* ((neg (alexandria:xor (< num 0) (and display-factor
							 (< display-factor 0))))
		     (dn (abs (/ num (or display-factor 1))))
		     (stem (if decimal-places
			       (if (and (integerp dn) decimal-replacement)
				   (format nil "~d.~a" dn decimal-replacement)
				   (format nil "~,vf" decimal-places dn))
			       (princ-number dn)))
		     (dot (position #\. stem))
		     (len (length stem)))
		(when (and min-integer-digits
			   (< (or dot len) min-integer-digits))
		  (let ((delta (- min-integer-digits (or dot len))))
		    (setf stem (concatenate 'string (make-string delta :initial-element #\0) stem)
			  dot (+ dot delta)
			  len (+ len delta))))
		(when dot
		  (setf (char stem dot) (decimal-separator locale)))
		(concatenate 'string
			     (if neg "-" "")
			     (if number-grouping
				 (apply #'concatenate 'string
					(nreverse
					 (cons (subseq stem (or dot len))
					       (iter (with ld = (or dot len))
						     (with gl = (grouping-count locale))
						     (with start = (mod ld gl))
						     (for i from start below ld by gl)
						     (unless (zerop i)
						       (when (= i start)
							 (collect (subseq stem 0 i)))
						       (collect (string (grouping-separator locale)) at beginning))
						     (collect (subseq stem i (+ i gl)) at beginning)))))
				 stem)))))))

(defmethod format-data ((style number-currency-style) data)
  (with-slots (format locale) style
    (apply #'concatenate 'string
	   (iter (for (kw val) on format by #'cddr)
		 (collect (ecase kw
			    (:symbol val)
			    (:text val)
			    (:number (format-number data val locale))))))))

(defmethod format-data ((style number-date-time-style) data)
  (with-slots (format) style
    (multiple-value-bind (nsec ss mm hh day mon year dow)
	(local-time:decode-timestamp data)
      (declare (ignore nsec))
      (apply #'concatenate 'string
	     (iter (for val in format)
		   (if (stringp val)
		       (collect val)
		       (ecase val
			 ((:short-quarter :long-quarter) (format nil "Q~d" (1+ (floor mon 3))))
			 (:long-year (format nil "~4,'0d" year))
			 (:short-year (format nil "~2,'0d" (mod year 100)))
			 (:long-month (format nil "~2,'0d" mon))
			 (:short-month (format nil "~d" mon))
			 (:long-day (format nil "~2,'0d" day))
			 (:short-day (format nil "~d" day))
			 (:long-hours (format nil "~2,'0d" #1=(if (find :am-pm format)
								 (1+ (mod (1- hh) 12))
								 hh)))
			 (:short-hours (format nil "~d" #1#))
			 (:long-minutes (format nil "~2,'0d" mm))
			 (:short-minutes (format nil "~d" mm))
			 (:long-seconds (format nil "~2,'0d" ss))
			 (:short-seconds (format nil "~d" ss))
			 (:am-pm (if (>= 12 hh) "PM" "AM"))
			 (:era (warn "era formatting not supported") "")
			 (:day-of-week (warn "day-of-week formatting not properly supported") (format nil "~d" dow))
			 (:week-of-year (warn "week-of-year formatting not supported") ""))))))))

(defmethod format-data ((style number-number-style) data)
  (with-slots (format prefix suffix locale) style
    (concatenate 'string
		 (or prefix "")
		 (format-number data format locale)
		 (or suffix ""))))

(defmethod format-data ((style number-percentage-style) data)
  (with-slots (format prefix suffix locale) style
    (concatenate 'string
		 (or prefix "")
		 (format-number (if (realp data) (* 100 data) data) format locale)
		 (or suffix ""))))

(defmethod format-data ((style number-text-style) data)
  (with-slots (prefix suffix) style
    (concatenate 'string
		 (or prefix "")
		 data
		 (or suffix ""))))
