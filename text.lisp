(in-package :clods-export)

(defparameter *line-types* '(:none :single :double)
  "The valid ODS line types.")

(defparameter *line-styles* '(:none :solid :dotted :dash :long-dash :dot-dash :dot-dot-dash :wave)
  "The valid ODS line styles.")

(defparameter *line-widths* '(:auto :normal :bold :thin :medium :thick)
  "The valid line widths. Alternatively, a standard CSS length spec can be used (e.g. \"2pt\" or \"0.1cm\").")

(defparameter *line-modes* '(:continuous :skip-white-space)
  "The valid line modes.")

(defparameter *text-transforms* '(:none :lowercase :uppercase :capitalize)
  "The valid text transformations.")

(defparameter *script-types* '(:latin :asian :complex :ignore)
  "The valid script types.")

(defparameter *font-pitches* '(:fixed :variable)
  "The list of font pitches supported by ODS documents.")

(defparameter *font-reliefs* '(:none :embossed :engraved)
  "The valid font relief modes.")

(defun validate-font-weight (x)
  "Check that the given datum is a valid font weight."
  (or (and (keywordp x)
	   (member x *font-weights*))
      (and (integerp x)
	   (zerop (mod x 100))
	   (<= 100 x 900))))

(defparameter *text-properties*
  (list (list :font-variant *ns-fo* *font-variants*)
	(list :text-transform *ns-fo* *text-transforms*)
	(list :color *ns-fo* #'valid-color)
	(list :use-window-font-color *ns-style* 'boolean)
	(list :text-outline *ns-style* 'boolean)
	(list :text-line-through-type *ns-style* *line-types*)
	(list :text-line-through-style *ns-style* *line-styles*)
	(list :text-line-through-width *ns-style* *line-widths*)
	(list :text-line-through-color *ns-style* #'(lambda (x) (or (eq x :font-color) (valid-color x))))
	(list :text-line-through-mode *ns-style* *line-modes*)
	(list :text-line-through-text *ns-style* 'string)
	(list :text-line-through-text-style *ns-style* 'string)
	(list :text-position *ns-style* '(:super :sub))
	(list :font-name *ns-style* 'string)
	(list :font-name-asian *ns-style* 'string)
	(list :font-name-complex *ns-style* 'string)
	(list :font-family *ns-style* 'string)
	(list :font-family-asian *ns-style* 'string)
	(list :font-family-complex *ns-style* 'string)
	(list :font-family-generic *ns-style* *font-generic-families*)
	(list :font-family-generic-asian *ns-style* *font-generic-families*)
	(list :font-family-generic-complex *ns-style* *font-generic-families*)
	(list :font-style-name *ns-style* 'string)
	(list :font-style-name-asian *ns-style* 'string)
	(list :font-style-name-complex *ns-style* 'string)
	(list :font-pitch *ns-style* *font-pitches*)
	(list :font-pitch-asian *ns-style* *font-pitches*)
	(list :font-pitch-complex *ns-style* *font-pitches*)
	(list :font-size *ns-fo* 'string)
	(list :font-size-asian *ns-style* 'string)
	(list :font-size-complex *ns-style* 'string)
	(list :font-size-rel *ns-style* 'string)
	(list :font-size-rel-asian *ns-style* 'string)
	(list :font-size-rel-complex *ns-style* 'string)
	(list :script-type *ns-style* *script-types*)
	(list :letter-spacing *ns-fo* 'string)
	(list :font-style *ns-fo* *font-styles*)
	(list :font-style-asian *ns-style* *font-styles*)
	(list :font-style-complex *ns-style* *font-styles*)
	(list :font-relief *ns-style* *font-reliefs*)
	(list :text-shadow *ns-fo* 'string)
	(list :text-underline-type *ns-style* *line-types*)
	(list :text-underline-style *ns-style* *line-styles*)
	(list :text-underline-width *ns-style* *line-widths*)
	(list :text-underline-color *ns-style* #'(lambda (x) (or (eq x :font-color) (valid-color x))))
	(list :text-underline-mode *ns-style* *line-modes*)
	(list :text-overline-type *ns-style* *line-types*)
	(list :text-overline-style *ns-style* *line-styles*)
	(list :text-overline-width *ns-style* *line-widths*)
	(list :text-overline-color *ns-style* #'(lambda (x) (or (eq x :font-color) (valid-color x))))
	(list :text-overline-mode *ns-style* *line-modes*)
	(list :font-weight *ns-fo* #'validate-font-weight)
	(list :font-weight-asian *ns-style* #'validate-font-weight)
	(list :font-weight-complex *ns-style* #'validate-font-weight)
	(list :letter-kerning *ns-style* 'boolean)
	(list :text-blinking *ns-style* 'boolean)
	(list :text-combine *ns-style* '(:none :letters :lines))
	(list :text-emphasize *ns-style* '((:none :accent :dot :circle :disc) (:above :below)))
	(list :text-scale *ns-style* 'string)
	(list :text-rotation-angle *ns-style* 'string)
	(list :text-rotation-scale *ns-style* '(:fixed :line-height))
	(list :hyphenate *ns-fo* 'boolean)
	(list :hyphenation-remain-char-count *ns-fo* 'integer)
	(list :hyphenation-push-char-count *ns-fo* 'integer))
  "The set of text properties supported, with the validation rules for their content.")

(defun write-text-properties (props)
  "Write a list containing text properties into the XML document."
  (flet ((any-value (x)
	   (etypecase x
	     (string x)
	     (symbol (string-downcase x))
	     (number (princ-number x)))))
    (with-tag ((*ns-style* "text-properties"))
      (loop for (kw data) on props by #'cddr
	    for (nil ns constr) = (find kw *text-properties* :key #'first)
	    do (check-type kw keyword)
	    unless ns do (error "unknown text property: ~s" kw)
	    do (attr (ns (string-downcase kw))
		     (cond ((eq 'string constr)
			    (check-type data string)
			    data)
			   ((eq 'integer constr)
			    (check-type data integer)
			    (princ-to-string data))
			   ((eq 'boolean constr)
			    (if data "true" "false"))
			   ((functionp constr)
			    (unless (funcall constr data)
			      #1=(error "invalid value ~s for text property ~s" data kw))
			    (any-value data))
			   ((and (listp constr) (listp (first constr)))
			    (check-type data list)
			    (apply #'concatenate 'string
				   (iter (for i in data)
					 (for c in constr)
					 (unless (find i c)
					   #1#)
					 (unless (eq i (first data))
					   (collect " "))
					 (collect (any-value i)))))
			   ((listp constr)
			    (unless (find data constr)
			      #1#)
			    (any-value data))))))))
