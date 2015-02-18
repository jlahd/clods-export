(in-package :cl-user)
(require "clods-export")
(require "local-time")

(defun write-example-ods (path)
  (let ((text-props-normal '(:color "#000000" :font-name "Libration Sans" :font-size "12pt"))
	(text-props-title '(:color "#4040c0" :font-name "Times New Roman" :font-size "24pt" :font-weight :bold))
	(text-props-header '(:color "#000000" :font-name "Libration Sans" :font-size "12pt" :font-weight :bold))
	(text-props-link '(:color "#0563c1" :font-name "Libration Sans" :font-size "12pt"
			   :text-underline-style :solid :text-underline-type :single)))
  (clods:with-spreadsheet (path :generator "write-example-ods" :creator "Test User")
    (clods:using-fonts ()
      ;; fonts
      (clods:font "Liberation Sans" :family "Liberation Sans" :family-generic :swiss)
      (clods:font "Times New Roman" :family "Times New Roman" :family-generic :roman))

    (clods:using-styles (:locale (clods:make-locale "FI" #\space 3 #\,))
      ;; number formats
      (clods:number-text-style "n-text")
      (clods:number-number-style "n-int" '(:min-integer-digits 1 :decimal-places 0))
      (clods:number-number-style "n-dec" '(:min-integer-digits 1 :decimal-places 2))
      (clods:number-number-style "n-frac" '(:min-integer-digits 0 :min-numerator-digits 1 :min-denominator-digits 3))
      (clods:number-number-style "n-weight" '(:min-integer-digits 1 :min-exponent-digits 1 :decimal-places 3) :suffix " kg")
      (clods:number-currency-style "n-price" '(:number (:min-integer-digits 1 :decimal-places 2)
					       :text " "
					       :symbol "€"))
      (clods:number-date-style "n-date" '(:long-day "." :long-month "." :long-year))

      ;; cell styles
      (clods:cell-style "ce-normal" nil text-props-normal :data-style "n-text")
      (clods:cell-style "ce-title" "ce-normal" text-props-title :horizontal-align :center)
      (clods:cell-style "ce-header" "ce-normal" text-props-header :border-bottom '(:thin :solid "#000000"))
      (clods:cell-style "ce-header-r" "ce-normal" text-props-header :border-bottom '(:thin :solid "#000000") :horizontal-align :end)
      (clods:cell-style "ce-header-group" "ce-normal" text-props-header :horizontal-align :center)
      (clods:cell-style "ce-link" "ce-normal" text-props-link)
      (clods:cell-style "ce-int" "ce-normal" text-props-normal :data-style "n-int")
      (clods:cell-style "ce-dec" "ce-normal" text-props-normal :data-style "n-dec")
      (clods:cell-style "ce-frac" "ce-normal" text-props-normal :data-style "n-frac")
      (clods:cell-style "ce-weight" "ce-normal" text-props-normal :data-style "n-weight")
      (clods:cell-style "ce-price" "ce-normal" text-props-normal :data-style "n-price")
      (clods:cell-style "ce-date" "ce-normal" text-props-normal :data-style "n-date")

      ;; column styles
      (clods:column-style "co-id" nil :width "1.0cm")
      (clods:column-style "co-name" nil :width "5.0cm")
      (clods:column-style "co-number" nil :width "3.0cm")
      (clods:column-style "co-date" nil :width "4.0cm")

      ;; row styles
      (clods:row-style "ro-normal" nil :height "18pt" :use-optimal-height t)
      (clods:row-style "ro-title" nil :height "28pt" :use-optimal-height t))

    (clods:with-body ()
      (clods:with-table ("Products")
	;; first, define the columns
	(clods:with-header-columns ()
	  (clods:column :style "co-id" :cell-style "ce-int"))
	(clods:column :style "co-name" :cell-style "ce-normal" :repeat 2)
	(clods:column :style "co-number" :cell-style "ce-int")
	(clods:column :style "co-number" :cell-style "ce-frac")
	(clods:column :style "co-number" :cell-style "ce-weight")
	(clods:column :style "co-number" :cell-style "ce-price" :repeat 2)
	(clods:column :style "co-date" :cell-style "ce-date")

	;; then, add the data row-by-row, starting with headers
	(clods:with-header-rows ()
	  (clods:with-row (:repeat 2)) ; two empty rows at top
	  (clods:with-row (:style "ro-title")
	    (clods:cell "Product listing" :style "ce-title" :span-columns 9))
	  (clods:with-row ())
	  (clods:with-row (:style "ro-normal")
	    (clods:cells nil nil nil nil nil nil)
	    (clods:cell "Price" :style "ce-header-group" :span-columns 2))
	  (clods:with-row (:style "ro-normal" :cell-style "ce-header-r")
	    (clods:cell "#")
	    (clods:cell "Title" :style "ce-header")
	    (clods:cell "Note" :style "ce-header")
	    (clods:cells "Quantity" "Rating" "Weight" "VAT 0%" "VAT 24%" "Availability")))

	(let ((products '(("Product one" nil 42 123/456 12 82.10 "2015-01-01")
			  ("Product two" "Fragile" 2 0.88 13840.11d0 11.55 "2015-02-01")
			  ("Product three" "Out of stock" -17 #.pi 0.0000077 191.91 "2015-05-01"))))
	  (loop for (title note quantity rating weight price availability) in products
		for id from 1
		do (clods:with-row (:style "ro-normal")
		     (clods:cells id title note quantity rating weight
				  price (* 1.24 price) (local-time:parse-timestring availability)))))

	(clods:with-row ())
	(clods:with-row (:style "ro-normal")
	  (clods:cell nil)
	  (clods:cell "Generated with clods-export" :style "ce-link" :span-columns 2
		      :link "https://github.com/jlahd/clods-export")))))))
