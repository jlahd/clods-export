(in-package :clods-export)

(defclass locale ()
  ((country :initform nil :initarg :country :reader locale-country)
   (grouping-separator :initform nil :initarg :grouping-separator :reader grouping-separator)
   (grouping-count :initform nil :initarg :grouping-count :reader grouping-count)
   (decimal-separator :initform #\. :initarg :decimal-separator :reader decimal-separator)))

(defparameter *default-locale* (make-instance 'locale))

(defun make-locale (country-code grouping-separator grouping-length decimal-separator)
  "Make a new locale with the given ISO 3166 country code, number
grouping separator, number grouping length, and decimal separator."
  (make-instance 'locale
		 :country country-code
		 :grouping-separator grouping-separator
		 :grouping-count grouping-length
		 :decimal-separator decimal-separator))
