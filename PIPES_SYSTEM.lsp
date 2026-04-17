;;; =============================================================
;;; PIPES_SYSTEM.lsp  v3.1
;;; HVAC refrigerant system analyzer for AutoCAD
;;;
;;; Select area -> finds ODU/IDU/Refnets/Pipes -> CSV
;;; Pipes detected from TEXT labels, no line tracing.
;;;
;;; Command: PIPES_SYSTEM
;;; =============================================================

(vl-load-com)

(setq *PS:MRAD* 5000.0)  ; MTEXT search radius (mm)
(setq *PS:DRAD* 1000.0)  ; diameter label search radius (mm)

;;; ================================================================
;;;  String utilities
;;; ================================================================

(defun ps:trim (s)
  (if (null s) (setq s ""))
  (while (and (> (strlen s) 0) (= (substr s 1 1) " "))
    (setq s (substr s 2)))
  (while (and (> (strlen s) 0) (= (substr s (strlen s) 1) " "))
    (setq s (substr s 1 (1- (strlen s)))))
  s)

(defun ps:sw (s pfx)
  (and (>= (strlen s) (strlen pfx))
       (= (strcase (substr s 1 (strlen pfx))) (strcase pfx))))

;;; Parse MTEXT: strip RTF, split by \P and real newlines
(defun ps:mtext->lines (raw / lines cur i len ch code)
  (setq raw (if raw raw "")
        lines nil  cur ""  i 1  len (strlen raw))
  (while (<= i len)
    (setq ch (substr raw i 1)
          code (ascii ch))
    (cond
      ;; Real newline chars (10=LF, 13=CR)
      ((or (= code 10) (= code 13))
       (if (> (strlen (ps:trim cur)) 0)
         (setq lines (append lines (list (ps:trim cur)))))
       (setq cur ""  i (1+ i)))
      ;; \P or \n paragraph break
      ((and (= ch "\\") (<= (1+ i) len)
            (member (strcase (substr raw (1+ i) 1)) '("P" "N")))
       (if (> (strlen (ps:trim cur)) 0)
         (setq lines (append lines (list (ps:trim cur)))))
       (setq cur ""  i (+ i 2)))
      ;; RTF code \X...;
      ((and (= ch "\\") (<= (1+ i) len)
            (not (member (strcase (substr raw (1+ i) 1)) '("P" "N" "~"))))
       (setq i (+ i 2))
       (while (and (<= i len) (not (= (substr raw i 1) ";")))
         (setq i (1+ i)))
       (setq i (1+ i)))
      ;; RTF braces
      ((member ch '("{" "}")) (setq i (1+ i)))
      ;; Normal char
      (t (setq cur (strcat cur ch)) (setq i (1+ i)))))
  (if (> (strlen (ps:trim cur)) 0)
    (setq lines (append lines (list (ps:trim cur)))))
  lines)

;;; Number -> string with comma (RU Excel)
(defun ps:n2s (n)
  (vl-string-subst "," "." (rtos n 2 2)))

;;; ================================================================
;;;  AutoCAD helpers
;;; ================================================================

(defun ps:etype (e) (cdr (assoc 0  (entget e))))
(defun ps:ipt   (e) (cdr (assoc 10 (entget e))))

(defun ps:d2 (a b)
  (sqrt (+ (expt (- (car a) (car b)) 2)
           (expt (- (cadr a) (cadr b)) 2))))

;;; ================================================================
;;;  MTEXT parsing
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
  (setq lines (ps:mtext->lines (cdr (assoc 1 (entget e))))
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
;;;  Block attributes + IDU model
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

;;; Extract clean model: only alphanumeric chars
(defun ps:clean-model (s / i ch result)
  (setq result ""  i 1)
  (while (<= i (strlen s))
    (setq ch (substr s i 1))
    (if (or (and (>= (ascii ch) 48) (<= (ascii ch) 57))
            (and (>= (ascii ch) 65) (<= (ascii ch) 90))
            (and (>= (ascii ch) 97) (<= (ascii ch) 122)))
      (setq result (strcat result ch))
      (if (> (strlen result) 0) (setq i (1+ (strlen s)))))
    (setq i (1+ i)))
  result)

;;; IDU model: attribute ARNU* -> nearby MTEXT ARNU* -> block name
(defun ps:vb-model (e bname / model me lines pos sub)
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
                  (setq model (ps:clean-model (substr line (1+ pos))))))))))))
  (if (and model (> (strlen model) 0)) model bname))

;;; ================================================================
;;;  Nearest ODU system
;;; ================================================================

(defun ps:nearest-sys (pt systems / d best-d best-name)
  (setq best-d 1e10  best-name nil)
  (foreach sys systems
    (setq d (ps:d2 pt (cadddr sys)))
    (if (< d best-d)
      (setq best-d d  best-name (car sys))))
  best-name)

;;; ================================================================
;;;  Find diameter label nearest to a point (contains ":", no "/")
;;; ================================================================

(defun ps:find-diam-near (cen rad / ss i e ed txt pt d best-d best)
  (setq best nil  best-d 1e10)
  (setq ss (ssget "_C"
             (list (- (car cen) rad) (- (cadr cen) rad) -9e9)
             (list (+ (car cen) rad) (+ (cadr cen) rad)  9e9)))
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
                     (null (vl-string-search "kW" txt))
                     (null (vl-string-search "kg" txt)))
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
;;;  Accumulator + CSV
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

;;; Helper: update sys-data-map for a system
(defun ps:map-accum (smap sname key num)
  (mapcar (function (lambda (sd)
    (if (= (car sd) sname)
      (list sname (ps:accum key num (cadr sd)))
      sd)))
    smap))

;;; ================================================================
;;;  Main command: PIPES_SYSTEM
;;; ================================================================

(defun C:PIPES_SYSTEM (/ ss i e ed etype bname ipt me parsed
                          sys-name model systems sys-data-map
                          nearest imodel key txt pt
                          diam lenval results all-keys fname)
  (vl-load-com)
  (princ "\n=== PIPES_SYSTEM v3.1 ===")
  (princ "\nSelect area with systems, then Enter: ")

  (setq ss (ssget))
  (if (null ss)
    (progn (princ "\nCancelled.") (exit)))
  (princ (strcat "\n" (itoa (sslength ss)) " objects selected."))

  ;; ---- Pass 1: find ODU blocks ----
  (setq systems nil  i 0)
  (repeat (sslength ss)
    (setq e  (ssname ss i)
          ed (entget e))
    (if (and (= (cdr (assoc 0 ed)) "INSERT")
             (ps:sw (cdr (assoc 2 ed)) "LATS_ODU"))
      (progn
        (setq ipt (cdr (assoc 10 ed)))
        (setq me (ps:mtext-near ipt *PS:MRAD* "ID"))
        (if me
          (progn
            (setq parsed   (ps:parse-mtext me))
            (setq sys-name (ps:find-val "ID" parsed))
            (setq model    (ps:find-val-sw "ARUN" parsed)))
          (setq sys-name nil  model nil))
        (if (null sys-name) (setq sys-name (cdr (assoc 2 ed))))
        (if (null model)    (setq model    (cdr (assoc 2 ed))))
        ;; (name model ent point)
        (setq systems (append systems
                (list (list sys-name model e ipt))))))
    (setq i (1+ i)))

  (if (null systems)
    (progn (alert "No LATS_ODU* in selection!") (exit)))
  (princ (strcat "\n" (itoa (length systems)) " systems."))

  ;; Init data map
  (setq sys-data-map nil)
  (foreach sys systems
    (setq sys-data-map
      (append sys-data-map
        (list (list (car sys)
                (list (cons (strcat "ODU " (cadr sys)) 1)))))))

  ;; ---- Pass 2: classify blocks + text ----
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
             ipt   (cdr (assoc 10 ed))
             nearest (ps:nearest-sys ipt systems))
       (if nearest
         (progn
           (setq imodel (ps:vb-model e bname))
           (setq sys-data-map
             (ps:map-accum sys-data-map nearest
               (strcat "IDU " imodel) 1)))))

      ;; Refnet blocks
      ((and (= etype "INSERT")
            (ps:sw (cdr (assoc 2 ed)) "LATS_SPLT"))
       (setq bname (cdr (assoc 2 ed))
             ipt   (cdr (assoc 10 ed))
             nearest (ps:nearest-sys ipt systems))
       (if nearest
         (setq sys-data-map
           (ps:map-accum sys-data-map nearest
             (strcat "Refnet " bname) 1))))

      ;; TEXT/MTEXT with "/" -> pipe length label
      ((and (member etype '("TEXT" "MTEXT"))
            (setq txt (cdr (assoc 1 ed)))
            (vl-string-search "/" txt)
            (null (vl-string-search "kW" txt))
            (null (vl-string-search "kg" txt)))
       (setq pt  (cdr (assoc 10 ed))
             nearest (ps:nearest-sys pt systems))
       (if nearest
         (progn
           ;; Find diameter label nearby
           (setq diam (ps:find-diam-near pt *PS:DRAD*))
           (setq lenval (ps:extract-len txt))
           (if (and diam (> lenval 0.0))
             (setq sys-data-map
               (ps:map-accum sys-data-map nearest
                 (strcat "Pipe " diam " m") lenval)))))))

    (setq i (1+ i)))

  ;; ---- Results ----
  (setq results sys-data-map)
  (foreach sys results
    (princ (strcat "\n  " (car sys) ": "
                   (itoa (length (cadr sys))) " items")))

  (setq all-keys (ps:all-keys results))
  (setq fname (getfiled "Save CSV" (getvar "DWGPREFIX") "csv" 1))
  (if (null fname)
    (progn (princ "\nCancelled.") (exit)))

  (if (ps:write-csv fname results all-keys)
    (progn
      (princ (strcat "\nDone! " fname))
      (startapp "explorer" fname))
    (alert "File write error!"))

  (princ))

;;; ================================================================
(princ "\nPIPES_SYSTEM v3.1 loaded. Command: PIPES_SYSTEM")
(princ)
