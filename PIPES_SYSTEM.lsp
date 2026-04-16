;;; =============================================================
;;; PIPES_SYSTEM.lsp  v3.0
;;; HVAC refrigerant system analyzer for AutoCAD
;;;
;;; User selects area -> script finds:
;;;   ODU  (LATS_ODU_*)    - outdoor units
;;;   IDU  (LATS_IDU_*)    - indoor units (model from MTEXT ARNU*)
;;;   Refnets (LATS_SPLT_*) - branch splitters
;;;   Pipes - lengths by diameter from text labels
;;;
;;; Groups by system (nearest ODU). Exports CSV.
;;; Command: PIPES_SYSTEM
;;; =============================================================

(vl-load-com)

;;; ---- Settings ----
(setq *PS:MRAD* 5000.0)  ; MTEXT search radius (mm)
(setq *PS:PRAD* 1000.0)  ; pipe label search radius (mm)

;;; ================================================================
;;;  String utilities
;;; ================================================================

(defun ps:trim (s / i)
  (if (null s) (setq s ""))
  (while (and (> (strlen s) 0) (= (substr s 1 1) " "))
    (setq s (substr s 2)))
  (while (and (> (strlen s) 0) (= (substr s (strlen s) 1) " "))
    (setq s (substr s 1 (1- (strlen s)))))
  s)

(defun ps:sw (s pfx)
  (and (>= (strlen s) (strlen pfx))
       (= (strcase (substr s 1 (strlen pfx))) (strcase pfx))))

;;; Parse MTEXT: strip RTF codes, split by \P
(defun ps:mtext->lines (raw / lines cur i len ch)
  (setq raw   (if raw raw "")
        lines nil  cur ""  i 1  len (strlen raw))
  (while (<= i len)
    (setq ch (substr raw i 1))
    (cond
      ((and (= ch "\\") (<= (1+ i) len)
            (member (strcase (substr raw (1+ i) 1)) '("P" "N")))
       (if (> (strlen (ps:trim cur)) 0)
         (setq lines (append lines (list (ps:trim cur)))))
       (setq cur "" i (+ i 2)))
      ((and (= ch "\\") (<= (1+ i) len)
            (not (member (strcase (substr raw (1+ i) 1)) '("P" "N" "~"))))
       (setq i (+ i 2))
       (while (and (<= i len) (not (= (substr raw i 1) ";")))
         (setq i (1+ i)))
       (setq i (1+ i)))
      ((member ch '("{" "}")) (setq i (1+ i)))
      (t (setq cur (strcat cur ch)) (setq i (1+ i)))))
  (if (> (strlen (ps:trim cur)) 0)
    (setq lines (append lines (list (ps:trim cur)))))
  lines)

;;; Number to string with comma decimal (RU Excel)
(defun ps:n2s (n)
  (vl-string-subst "," "." (rtos n 2 2)))

;;; ================================================================
;;;  AutoCAD helpers
;;; ================================================================

(defun ps:etype (e) (cdr (assoc 0  (entget e))))
(defun ps:bname (e) (cdr (assoc 2  (entget e))))
(defun ps:ipt   (e) (cdr (assoc 10 (entget e))))
(defun ps:txt   (e) (cdr (assoc 1  (entget e))))

;;; 2D distance
(defun ps:d2 (a b)
  (sqrt (+ (expt (- (car  a) (car  b)) 2)
           (expt (- (cadr a) (cadr b)) 2))))

;;; ================================================================
;;;  MTEXT search and parsing
;;; ================================================================

(defun ps:mtext-near (cen rad key / ss i e ed txt pt d best-d best-e)
  (setq best-d 1e10  best-e nil)
  (setq ss (ssget "_C"
             (list (- (car cen) rad) (- (cadr cen) rad) -9e9)
             (list (+ (car cen) rad) (+ (cadr cen) rad)  9e9)))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i)  ed (entget e))
        (if (= (cdr (assoc 0 ed)) "MTEXT")
          (progn
            (setq txt (cdr (assoc 1 ed)))
            (if (and txt (vl-string-search key txt))
              (progn
                (setq pt (cdr (assoc 10 ed))
                      d  (ps:d2 cen pt))
                (if (< d best-d) (setq best-d d  best-e e))))))
        (setq i (1+ i)))))
  best-e)

(defun ps:parse-mtext (e / lines line col k v result)
  (setq lines  (ps:mtext->lines (cdr (assoc 1 (entget e))))
        result nil)
  (foreach line lines
    (setq col (vl-string-search ":" line))
    (if col
      (progn
        (setq k (ps:trim (substr line 1 col))
              v (ps:trim (substr line (+ col 2))))
        (setq result (append result (list (cons k v)))))))
  result)

(defun ps:find-val (frag alist / found)
  (setq found nil)
  (foreach p alist
    (if (and (null found)
             (vl-string-search (strcase frag) (strcase (car p))))
      (setq found (cdr p))))
  found)

(defun ps:find-val-sw (prefix alist / found)
  (setq found nil)
  (foreach p alist
    (if (and (null found) (ps:sw (cdr p) prefix))
      (setq found (cdr p))))
  found)

;;; ================================================================
;;;  Block attributes
;;; ================================================================

(defun ps:get-attrs (e / sub ed result)
  (setq result nil  sub (entnext e))
  (while (and sub (= (ps:etype sub) "ATTRIB"))
    (setq ed (entget sub))
    (setq result (append result
                   (list (cons (cdr (assoc 2 ed))
                               (cdr (assoc 1 ed))))))
    (setq sub (entnext sub)))
  result)

(defun ps:attr-sw (e prefix / found)
  (setq found nil)
  (foreach p (ps:get-attrs e)
    (if (and (null found) (ps:sw (cdr p) prefix))
      (setq found (cdr p))))
  found)

;;; IDU model: attribute ARNU* -> nearby MTEXT -> block name
(defun ps:vb-model (e bname / model me lines pos sub sp)
  (setq model (ps:attr-sw e "ARNU"))
  (if (null model)
    (progn
      (setq me (ps:mtext-near (ps:ipt e) *PS:MRAD* "ARNU"))
      (if me
        (progn
          (setq lines (ps:mtext->lines (cdr (assoc 1 (entget me)))))
          (foreach line lines
            (if (null model)
              (progn
                (setq pos (vl-string-search "ARNU" line))
                (if pos
                  (progn
                    (setq sub (substr line (1+ pos))
                          sp  (vl-string-search " " sub))
                    (setq model (if sp (substr sub 1 sp) sub)))))))))))
  (if (and model (> (strlen model) 0)) model bname))

;;; ================================================================
;;;  Find nearest ODU from a list of systems
;;; ================================================================

(defun ps:nearest-sys (pt systems / d best-d best-name)
  (setq best-d 1e10  best-name nil)
  (foreach sys systems
    (setq d (ps:d2 pt (cadddr sys)))
    (if (< d best-d)
      (setq best-d d  best-name (car sys))))
  best-name)

;;; ================================================================
;;;  Pipe label search (diameter above / length below a point)
;;; ================================================================

(defun ps:diam-above (cen rad / ss i e ed txt pt d best-d best)
  (setq best nil  best-d 1e10)
  (setq ss (ssget "_C"
             (list (- (car cen) rad) (cadr cen) -9e9)
             (list (+ (car cen) rad) (+ (cadr cen) rad) 9e9)))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i)  ed (entget e))
        (if (member (cdr (assoc 0 ed)) '("TEXT" "MTEXT"))
          (progn
            (setq txt (cdr (assoc 1 ed))
                  pt  (cdr (assoc 10 ed)))
            (if (and txt pt
                     (vl-string-search ":" txt)
                     (null (vl-string-search "/" txt))
                     (null (vl-string-search "ID" txt))
                     (> (cadr pt) (cadr cen)))
              (progn
                (setq d (ps:d2 cen pt))
                (if (< d best-d) (setq best-d d  best (ps:trim txt)))))))
        (setq i (1+ i)))))
  best)

(defun ps:len-below (cen rad / ss i e ed txt pt d best-d best)
  (setq best nil  best-d 1e10)
  (setq ss (ssget "_C"
             (list (- (car cen) rad) (- (cadr cen) rad) -9e9)
             (list (+ (car cen) rad) (cadr cen) 9e9)))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i)  ed (entget e))
        (if (member (cdr (assoc 0 ed)) '("TEXT" "MTEXT"))
          (progn
            (setq txt (cdr (assoc 1 ed))
                  pt  (cdr (assoc 10 ed)))
            (if (and txt pt (vl-string-search "/" txt))
              (progn
                (setq d (ps:d2 cen pt))
                (if (< d best-d) (setq best-d d  best (ps:trim txt)))))))
        (setq i (1+ i)))))
  best)

;;; Extract first number from "X,X / Y m(N)"
(defun ps:extract-len (s / pos)
  (if (null s) 0.0
    (progn
      (setq pos (vl-string-search "/" s))
      (if pos
        (atof (vl-string-subst "." "," (ps:trim (substr s 1 pos))))
        0.0))))

;;; ================================================================
;;;  Result accumulation + CSV
;;; ================================================================

(defun ps:accum (key num lst / p)
  (if (setq p (assoc key lst))
    (subst (cons key (+ (cdr p) num)) p lst)
    (append lst (list (cons key num)))))

(defun ps:val (key lst / p)
  (setq p (assoc key lst))
  (if p (cdr p) 0.0))

(defun ps:all-keys (results / keys)
  (setq keys nil)
  (foreach sys results
    (foreach p (cadr sys)
      (if (null (member (car p) keys))
        (setq keys (append keys (list (car p)))))))
  keys)

(defun ps:write-csv (path results all-keys / f hdr row total v)
  (setq f (open path "w"))
  (if (null f) nil
    (progn
      (setq hdr "Name")
      (foreach sys results (setq hdr (strcat hdr ";" (car sys))))
      (write-line (strcat hdr ";Total") f)
      (foreach key all-keys
        (setq row key  total 0.0)
        (foreach sys results
          (setq v (ps:val key (cadr sys)))
          (setq total (+ total v))
          (setq row (strcat row ";" (ps:n2s v))))
        (write-line (strcat row ";" (ps:n2s total)) f))
      (close f)
      T)))

;;; ================================================================
;;;  Main command: PIPES_SYSTEM
;;; ================================================================

(defun C:PIPES_SYSTEM (/ ss i e ed etype bname ipt me parsed
                          sys-name model systems sys-pts
                          nearest sys-data-map sys-data
                          imodel key txt pt diam lenstr lenval
                          results all-keys fname)
  (vl-load-com)
  (princ "\n=== PIPES_SYSTEM v3.0 ===")
  (princ "\nSelect area with systems, then Enter: ")

  ;; 1. User selects objects
  (setq ss (ssget))
  (if (null ss)
    (progn (princ "\nNo selection. Cancelled.") (exit)))

  (princ (strcat "\n" (itoa (sslength ss)) " objects selected."))

  ;; 2. First pass: find all ODU blocks -> systems
  (setq systems nil  i 0)
  (repeat (sslength ss)
    (setq e  (ssname ss i)
          ed (entget e))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq bname (cdr (assoc 2 ed)))
        (if (ps:sw bname "LATS_ODU")
          (progn
            (setq ipt (cdr (assoc 10 ed)))
            (setq me (ps:mtext-near ipt *PS:MRAD* "ID"))
            (if me
              (progn
                (setq parsed   (ps:parse-mtext me))
                (setq sys-name (ps:find-val "ID" parsed))
                (setq model    (ps:find-val-sw "ARUN" parsed)))
              (setq sys-name bname  model bname))
            (if (null sys-name) (setq sys-name bname))
            (if (null model)    (setq model    bname))
            ;; (sys-name model ent insertion-point)
            (setq systems (append systems
                    (list (list sys-name model e ipt))))))))
    (setq i (1+ i)))

  (if (null systems)
    (progn (alert "No LATS_ODU* blocks in selection!") (exit)))

  (princ (strcat "\n" (itoa (length systems)) " systems found."))

  ;; Init data map: ((sys-name . data-alist) ...)
  (setq sys-data-map nil)
  (foreach sys systems
    (setq sys-data-map
      (append sys-data-map
        (list (list (car sys)
                (list (cons (strcat "ODU " (cadr sys)) 1)))))))

  ;; 3. Second pass: classify all other objects -> nearest ODU
  (setq i 0)
  (repeat (sslength ss)
    (setq e     (ssname ss i)
          ed    (entget e)
          etype (cdr (assoc 0 ed)))
    (cond
      ;; IDU blocks
      ((and (= etype "INSERT")
            (ps:sw (cdr (assoc 2 ed)) "LATS_IDU"))
       (setq bname (cdr (assoc 2 ed))
             ipt   (cdr (assoc 10 ed)))
       (setq nearest (ps:nearest-sys ipt systems))
       (if nearest
         (progn
           (setq imodel (ps:vb-model e bname))
           (setq key (strcat "IDU " imodel))
           ;; Update data for this system
           (setq sys-data-map
             (mapcar (function (lambda (sd)
               (if (= (car sd) nearest)
                 (list nearest (ps:accum key 1 (cadr sd)))
                 sd)))
               sys-data-map)))))

      ;; Refnet blocks
      ((and (= etype "INSERT")
            (ps:sw (cdr (assoc 2 ed)) "LATS_SPLT"))
       (setq bname (cdr (assoc 2 ed))
             ipt   (cdr (assoc 10 ed)))
       (setq nearest (ps:nearest-sys ipt systems))
       (if nearest
         (setq sys-data-map
           (mapcar (function (lambda (sd)
             (if (= (car sd) nearest)
               (list nearest (ps:accum (strcat "Refnet " bname) 1 (cadr sd)))
               sd)))
             sys-data-map))))

      ;; Text with "/" -> pipe length label
      ((and (member etype '("TEXT" "MTEXT"))
            (progn (setq txt (cdr (assoc 1 ed))) nil)
            nil))

      ;; Lines -> find pipe labels near midpoint
      ((member etype '("LINE" "LWPOLYLINE" "POLYLINE"))
       (setq ipt (cdr (assoc 10 ed)))
       (if (= etype "LINE")
         (setq pt (list (/ (+ (car ipt) (car (cdr (assoc 11 ed)))) 2.0)
                        (/ (+ (cadr ipt) (cadr (cdr (assoc 11 ed)))) 2.0) 0.0))
         (setq pt ipt))
       (setq nearest (ps:nearest-sys pt systems))
       (if nearest
         (progn
           (setq diam   (ps:diam-above pt *PS:PRAD*))
           (setq lenstr (ps:len-below  pt *PS:PRAD*))
           (if (and diam lenstr)
             (progn
               (setq lenval (ps:extract-len lenstr))
               (if (> lenval 0.0)
                 (setq sys-data-map
                   (mapcar (function (lambda (sd)
                     (if (= (car sd) nearest)
                       (list nearest (ps:accum (strcat "Pipe " diam " m")
                                               lenval (cadr sd)))
                       sd)))
                     sys-data-map)))))))))

    (setq i (1+ i)))

  ;; 4. Build results
  (setq results sys-data-map)

  ;; Print summary
  (foreach sys results
    (princ (strcat "\n  " (car sys) ": "
                   (itoa (length (cadr sys))) " items")))

  ;; 5. Save CSV
  (setq all-keys (ps:all-keys results))
  (setq fname (getfiled "Save CSV report" (getvar "DWGPREFIX") "csv" 1))
  (if (null fname)
    (progn (princ "\nCancelled.") (exit)))

  (if (ps:write-csv fname results all-keys)
    (progn
      (princ (strcat "\nDone! " fname))
      (startapp "explorer" fname))
    (alert "File write error!"))

  (princ "\n=== PIPES_SYSTEM done ===")
  (princ))

;;; ================================================================
(princ "\nPIPES_SYSTEM v3.0 loaded. Command: PIPES_SYSTEM")
(princ)
