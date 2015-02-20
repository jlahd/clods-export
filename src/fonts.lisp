(in-package :clods-export)

(defparameter *font-generic-families* '(:roman :swiss :modern :decorative :script :system)
  "The list of generic font families supported by ODS documents.")

(defparameter *font-styles* '(:normal :italic :oblique)
  "The list of font styles supported by ODS documents.")

(defparameter *font-weights* '(:normal :bold :bolder :lighter)
  "The list of font weights supported by ODS documents. Additionally, numeric values 100-900 can be used.")

(defparameter *font-variants* '(:normal :small-caps)
  "The list of font variants supported by ODS documents.")

(defparameter *font-sizes* '(:xx-small :x-small :small :medium :large :x-large :xx-large :larger :smaller)
  "The list of font sizes supported by ODS documents.
Alternatively, a string value can be used specifying the font size in CSS format (e.g. \"18pt\" or \"1.2cm)\".")

(defparameter *font-stretches* '(:ultra-condensed :extra-condensed :condensed :semi-condensed
				 :normal
				 :semi-expanded :expanded :extra-expanded :ultra-expanded)
  "The list of font stretch values supported by ODS documents.")

(defmacro using-fonts (() &body body)
  "Specify the fonts to be used in the document."
  `(progn
     (unless (eq *sheet-state* :start)
       (error "with-fonts must appear as the first child of with-spreadsheet form"))
     (with-element* ("office" "font-face-decls")
       (let ((*sheet-state* :with-fonts))
	 ,@body))
     (setf *sheet-state* :fonts-defined)))

(defun font (name &key adornments family family-generic style weight size variant stretch)
  "Define a font with the given name and specs.
The font can be referenced in style definition using the given name."
  (unless (eq *sheet-state* :with-fonts)
    (error "font must be called in the context of a with-fonts form"))
  (with-element* ("style" "font-face")
    (attribute* "style" "name" name)
    (when adornments
      (check-type adornments string)
      (attribute* "svg" "font-adornments" adornments))
    (when family (attribute* "svg" "font-family" family))
    (when family-generic (attribute* "svg" "font-family-generic" (kw-to-string family-generic *font-generic-families*)))
    (when style (attribute* "svg" "font-style" (kw-to-string style *font-styles*)))
    (when weight (attribute* "svg" "font-weight"
		       (if (numberp weight)
			   (princ-to-string weight)
			   (kw-to-string weight *font-weights*))))
    (when variant (attribute* "svg" "font-variant" (kw-to-string variant *font-variants*)))
    (when size (attribute* "svg" "font-size"
		     (if (stringp size)
			 size
			 (kw-to-string size *font-sizes*))))
    (when stretch (attribute* "svg" "font-stretch"
			(kw-to-string stretch *font-stretches*)))))
