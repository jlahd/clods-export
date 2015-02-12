(in-package :zip)

(defun write-zipentry
    (z name data &key (file-write-date (file-write-date data)))
  (setf name (substitute #\/ #\\ name))
  (let* ((s (zipwriter-stream z))
         (header (make-local-header))
         (utf8-name (string-to-octets name (zipwriter-external-format z)))
         (entry (make-zipwriter-entry
                 :name name
                 :position (file-position s)
                 :header header)))
    (setf (file/signature header) #x04034b50)
    (setf (file/version-needed-to-extract header) 20) ;XXX ist das 2.0?
    (setf (file/flags header) 8)        ;bit 3: descriptor folgt nach daten
    (setf (file/method header) 8)
    (multiple-value-bind (s min h d m y)
        (decode-universal-time
         (or file-write-date (encode-universal-time 0 0 0 1 1 1980 0)))
      (setf (file/time header)
            (logior (ash h 11) (ash min 5) (ash s -1)))
      (setf (file/date header)
            (logior (ash (- y 1980) 9) (ash m 5) d)))
    (setf (file/compressed-size header) 0)
    (setf (file/size header) 0)
    (setf (file/name-length header) (length utf8-name))
    (setf (file/extra-length header) 0)
    (setf (zipwriter-tail z)
          (setf (cdr (zipwriter-tail z)) (cons entry nil)))
    (write-sequence header s)
    (write-sequence utf8-name s)
    (let ((descriptor (make-data-descriptor)))
      (multiple-value-bind (nin nout crc)
          (compress data s (zipwriter-compressor z))
        (setf (data/crc descriptor) crc)
        (setf (data/compressed-size descriptor) nout)
        (setf (data/size descriptor) nin)
        ;; record same values for central directory
        (setf (file/crc header) crc)
        (setf (file/compressed-size header) nout)
        (setf (file/size header) nin)) 
      (write-sequence descriptor s))
    name))
