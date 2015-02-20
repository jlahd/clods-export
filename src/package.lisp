(cl:defpackage :clods-export
  (:nicknames #:clods)
  (:use #:cl #:iterate #:cxml)
  (:export
   #:with-spreadsheet
   #:make-locale
   #:using-fonts #:font #:*font-generic-families* #:*font-styles* #:*font-weights* #:*font-variants* #:*font-sizes* #:*font-stretches*
   #:*line-types* #:*line-styles* #:*line-widths* #:*line-modes* #:*text-transforms* #:*script-types* #:*font-reliefs* #:*font-pitches*
   #:using-styles #:*text-properties*
   #:number-boolean-style #:number-date-style #:number-time-style #:number-text-style
   #:number-number-style #:number-percentage-style #:number-currency-style
   #:text-style #:table-style #:column-style #:row-style #:cell-style
   #:with-body #:with-table #:with-header-columns #:column #:with-header-rows #:with-row #:cell #:covered-cell #:cells))
