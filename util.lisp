(in-package :clods-export)

(defun kw-to-string (kw &optional allowed)
  "Convert a keyword into a lower-case string.
The optional allowed parameter can be specified to restrict the set of
accepted keywords (:kw1 :kw2 ...) and/or to specify different strings
for the keywords ((:kw1 \"kw1-string\") (:kw2 \"kw2-string\") ...).
A keyword not on the allowed keywords list will signal a type-error."
  (check-type kw keyword)
  (cond ((and allowed (listp (first allowed)))
	 (let ((match (find kw allowed :key #'first)))
	   (unless match
	     (error 'type-error :expected-type (cons 'member (mapcar #'first allowed)) :datum kw))
	   (second match)))
	((or (null allowed)
	     (find kw allowed))
	 (string-downcase kw))
	(t
	 (error 'type-error :expected-type (cons 'member allowed) :datum kw))))

(defun valid-color (c)
  "Validate a color specification.
A valid color is a six-digit hexadecimal number preceded by a #."
  (and (stringp c)
       (= (length c) 7)
       (char= #\# (char c 0))
       (ignore-errors (parse-integer c :start 1 :radix 16))
       t))

(defun princ-number (num)
  "princ-to-string a number, removing the trailing D0 in case of a double."
  (let* ((str (princ-to-string num))
	 (len (length str)))
    (if (and (> len 2)
	     (char= #\D (char-upcase (char str (- len 2))))
	     (char= #\0 (char str (- len 1))))
	(subseq str 0 (- len 2))
	str)))
