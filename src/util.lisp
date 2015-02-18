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
  "princ-to-string a number, enforcing a decimal, non-exponential form."
  (let* ((str (etypecase num
		(integer (princ-to-string num))
		(rational (format nil "~,100f" (coerce num 'double-float)))
		(real (format nil "~,100f" num))))
	 (dot (position #\. str)))
    (if dot
	(subseq str 0 (1+ (max (position-if-not (alexandria:curry #'char= #\0) str :from-end t)
			       (1+ dot))))
	str)))

(defun remove-nsec (time)
  "Remove the nanoseconds part from a local-time generated timestring,
e.g. \"2015-02-18T08:07:02.000000Z\" to \"2015-02-18T08:07:02Z\""
  (let ((zeros (search ".000000" time)))
    (if zeros
	(concatenate 'string (subseq time 0 zeros) (subseq time (+ zeros 7)))
	time)))

