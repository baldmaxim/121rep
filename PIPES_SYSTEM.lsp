;;; =============================================================
;;; PIPES_SYSTEM.lsp  v4.0
;;; HVAC system analyzer — connectivity via lines
;;;
;;; Command: PIPES_SYSTEM
;;; =============================================================

(vl-load-com)

(setq *PS:MRAD* 5000.0)  ; MTEXT search radius (mm)
(setq *PS:DRAD* 1000.0)  ; diameter label search radius (mm)
(setq *PS:BCON*  500.0)  ; block-to-line connection radius (mm)
(setq *PS:LTOL*   50.0)  ; line-to-line endpoint tolerance (mm)

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

(defun ps:mtext->lines (raw / lines cur i len ch code)
  (setq raw (if raw raw "")
        lines nil  cur ""  i 1  len (strlen raw))
  (while (<= i len)
    (setq ch (substr raw i 1)  code (ascii ch))
    (cond
      ((or (= code 10) (= code 13))
       (if (> (strlen (ps:trim cur)) 0)
         (setq lines (append lines (list (ps:trim cur)))))
       (setq cur ""  i (1+ i)))
      ((and (= ch "\\") (<= (1+ i) len)
            (member (strcase (substr raw (1+ i) 1)) '("P" "N")))
       (if (> (strlen (ps:trim cur)) 0)
         (setq lines (append lines (list (ps:trim cur)))))
       (setq cur ""  i (+ i 2)))
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

;;; Line/poly endpoints
(defun ps:line-pts (e / ed et pts sub)
  (setq ed (entget e)  et (cdr (assoc 0 ed)))
  (cond
    ((= et "LINE")
     (list (cdr (assoc 10 ed)) (cdr (assoc 11 ed))))
    ((= et "LWPOLYLINE")
     (setq pts nil)
     (foreach pair ed
       (if (= (car pair) 10)
         (setq pts (append pts (list (cdr pair))))))
     pts)
    ((= et "POLYLINE")
     (setq pts nil  sub (entnext e))
     (while (and sub (not (= (ps:etype sub) "SEQEND")))
       (if (= (ps:etype sub) "VERTEX")
         (setq pts (append pts (list (cdr (assoc 10 (entget sub)))))))
       (setq sub (entnext sub)))
     pts)
    (t nil)))

;;; Connection points for BFS:
;;;   Lines -> actual endpoints
;;;   Blocks -> all line endpoints within *PS:BCON* of insertion point
(defun ps:conn-pts (e all-lpts / ed et ipt pts)
  (setq ed (entget e)  et (cdr (assoc 0 ed)))
  (if (= et "INSERT")
    (progn
      (setq ipt (cdr (assoc 10 ed))  pts nil)
      (foreach p all-lpts
        (if (<= (ps:d2 ipt p) *PS:BCON*)
          (setq pts (append pts (list p)))))
      (if (null pts) (list ipt) pts))
    (ps:line-pts e)))

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
;;;  Block attributes + model detection
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

;;; IDU model
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

;;; Find nearest TEXT/MTEXT containing keyword
(defun ps:text-near (cen rad key / ss i e ed txt pt d best-d best)
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
            (if (and txt (vl-string-search key txt))
              (progn
                (setq d (ps:d2 cen pt))
                (if (< d best-d) (setq best-d d  best (ps:trim txt)))))))
        (setq i (1+ i)))))
  best)

;;; Refnet model from nearby TEXT ARBLN*
(defun ps:refnet-model (e bname / txt)
  (setq txt (ps:text-near (ps:ipt e) *PS:MRAD* "ARBLN"))
  (if (and txt (> (strlen txt) 0))
    (ps:clean-model txt)
    bname))

;;; ================================================================
;;;  BFS through connected entities
;;; ================================================================

;;; Precompute connection points for all pool entities
;;; Returns: ((entity . points-list) ...)
(defun ps:precompute-cp (pool all-lpts / result)
  (setq result nil)
  (foreach e pool
    (setq result (cons (cons e (ps:conn-pts e all-lpts)) result)))
  result)

;;; BFS from start-pts through precomputed pool
;;; Returns list of reached entities
(defun ps:bfs (start-pts pool-cp tol / queue visited result
                ent pts found cur-pts)
  (setq visited nil  result nil  queue nil)
  ;; Initial connections
  (foreach item pool-cp
    (setq ent (car item)  pts (cdr item)  found nil)
    (foreach sp start-pts
      (foreach ep pts
        (if (<= (ps:d2 sp ep) tol) (setq found T))))
    (if found
      (setq visited (cons ent visited)
            result  (cons ent result)
            queue   (cons pts queue))))
  ;; BFS loop
  (while queue
    (setq cur-pts (car queue)  queue (cdr queue))
    (foreach item pool-cp
      (setq ent (car item)  pts (cdr item))
      (if (not (member ent visited))
        (progn
          (setq found nil)
          (foreach p1 cur-pts
            (foreach p2 pts
              (if (<= (ps:d2 p1 p2) tol) (setq found T))))
          (if found
            (setq visited (cons ent visited)
                  result  (cons ent result)
                  queue   (cons pts queue)))))))
  result)

;;; ================================================================
;;;  Pipe labels
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

(defun ps:extract-len (s / pos)
  (if (null s) 0.0
    (progn
      (setq pos (vl-string-search "/" s))
      (if pos
        (atof (vl-string-subst "." "," (ps:trim (substr s 1 pos))))
        0.0))))

(defun ps:split-diams (s / pos)
  (setq pos (vl-string-search ":" s))
  (if pos
    (list (ps:trim (substr s 1 pos))
          (ps:trim (substr s (+ pos 2))))
    (list (ps:trim s))))

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

(defun ps:sort-keys (keys / odu idu ref pipe)
  (setq odu nil  idu nil  ref nil  pipe nil)
  (foreach k keys
    (cond
      ((ps:sw k "ODU")    (setq odu  (append odu  (list k))))
      ((ps:sw k "IDU")    (setq idu  (append idu  (list k))))
      ((ps:sw k "Refnet") (setq ref  (append ref  (list k))))
      (t                  (setq pipe (append pipe (list k))))))
  (append odu idu ref pipe))

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
;;;  Main command
;;; ================================================================

(defun C:PIPES_SYSTEM (/ ss i e ed etype bname ipt me parsed
                          sys-name model systems
                          pool all-lpts pool-cp odu-pts reached
                          sys-lines-map sys-data-map
                          imodel key txt pt diam lenval d
                          best-sys best-d line-sys
                          results all-keys fname)
  (vl-load-com)
  (princ "\n=== PIPES_SYSTEM v4.0 ===")
  (princ "\nSelect area, then Enter: ")

  (setq ss (ssget))
  (if (null ss)
    (progn (princ "\nCancelled.") (exit)))
  (princ (strcat "\n" (itoa (sslength ss)) " objects selected."))

  ;; ---- Pass 1: find ODU + build pool ----
  (setq systems nil  pool nil  i 0)
  (repeat (sslength ss)
    (setq e  (ssname ss i)
          ed (entget e)
          etype (cdr (assoc 0 ed)))
    (cond
      ;; Lines -> pool
      ((member etype '("LINE" "LWPOLYLINE" "POLYLINE"))
       (setq pool (cons e pool)))
      ;; Blocks
      ((= etype "INSERT")
       (setq bname (cdr (assoc 2 ed)))
       (cond
         ;; ODU -> systems
         ((ps:sw bname "LATS_ODU")
          (setq ipt (cdr (assoc 10 ed)))
          (setq me (ps:mtext-near ipt *PS:MRAD* "ID"))
          (if me
            (progn
              (setq parsed   (ps:parse-mtext me))
              (setq sys-name (ps:find-val "ID" parsed))
              (setq model    (ps:find-val-sw "ARUN" parsed)))
            (setq sys-name nil  model nil))
          (if (null sys-name) (setq sys-name bname))
          (if (null model)    (setq model bname))
          (setq systems (cons (list sys-name model e ipt) systems)))
         ;; IDU + Refnet -> pool
         ((or (ps:sw bname "LATS_IDU") (ps:sw bname "LATS_SPLT"))
          (setq pool (cons e pool))))))
    (setq i (1+ i)))

  (if (null systems)
    (progn (alert "No LATS_ODU* in selection!") (exit)))
  (princ (strcat "\n" (itoa (length systems)) " systems, "
                 (itoa (length pool)) " pool objects."))

  ;; ---- Collect all line endpoints + precompute ----
  (setq all-lpts nil)
  (foreach e pool
    (if (member (ps:etype e) '("LINE" "LWPOLYLINE" "POLYLINE"))
      (setq all-lpts (append all-lpts (ps:line-pts e)))))

  (princ (strcat "\n" (itoa (length all-lpts)) " line endpoints."))
  (setq pool-cp (ps:precompute-cp pool all-lpts))

  ;; ---- BFS from each ODU ----
  ;; sys-lines-map: ((sys-name . reached-entities) ...)
  (setq sys-lines-map nil)
  (setq sys-data-map nil)

  (foreach sys systems
    (setq sys-name (car sys)
          model    (cadr sys)
          ipt      (cadddr sys))
    (princ (strcat "\n  BFS: " sys-name "..."))

    ;; ODU connect points = line endpoints near ODU
    (setq odu-pts (ps:conn-pts (caddr sys) all-lpts))
    ;; BFS through pool
    (setq reached (ps:bfs odu-pts pool-cp *PS:LTOL*))
    (princ (strcat " " (itoa (length reached)) " connected."))

    (setq sys-lines-map (cons (cons sys-name reached) sys-lines-map))

    ;; Init system data with ODU
    (setq sys-data-map
      (cons (list sys-name (list (cons (strcat "ODU " model) 1)))
            sys-data-map)))

  ;; ---- Process reached entities per system ----
  (foreach slm sys-lines-map
    (setq sys-name (car slm))
    (foreach e (cdr slm)
      (setq etype (ps:etype e)
            bname (if (= etype "INSERT") (cdr (assoc 2 (entget e))) nil))
      (cond
        ;; IDU
        ((and bname (ps:sw bname "LATS_IDU"))
         (setq imodel (ps:vb-model e bname))
         (setq sys-data-map
           (mapcar (function (lambda (sd)
             (if (= (car sd) sys-name)
               (list sys-name (ps:accum (strcat "IDU " imodel) 1 (cadr sd)))
               sd)))
             sys-data-map)))
        ;; Refnet
        ((and bname (ps:sw bname "LATS_SPLT"))
         (setq imodel (ps:refnet-model e bname))
         (setq sys-data-map
           (mapcar (function (lambda (sd)
             (if (= (car sd) sys-name)
               (list sys-name (ps:accum (strcat "Refnet " imodel) 1 (cadr sd)))
               sd)))
             sys-data-map))))))

  ;; ---- Process text labels -> assign to system by nearest reached line ----
  (setq i 0)
  (repeat (sslength ss)
    (setq e  (ssname ss i)
          ed (entget e)
          etype (cdr (assoc 0 ed)))
    (if (and (member etype '("TEXT" "MTEXT"))
             (setq txt (cdr (assoc 1 ed)))
             (vl-string-search "/" txt)
             (null (vl-string-search "kW" txt))
             (null (vl-string-search "kg" txt)))
      (progn
        (setq pt (cdr (assoc 10 ed)))
        ;; Find which system has the nearest reached line
        (setq best-sys nil  best-d 1e10)
        (foreach slm sys-lines-map
          (foreach re (cdr slm)
            (if (member (ps:etype re) '("LINE" "LWPOLYLINE" "POLYLINE"))
              (progn
                (foreach lp (ps:line-pts re)
                  (setq d (ps:d2 pt lp))
                  (if (< d best-d)
                    (setq best-d d  best-sys (car slm))))))))
        ;; Accumulate pipe data
        (if best-sys
          (progn
            (setq diam (ps:find-diam-near pt *PS:DRAD*))
            (setq lenval (ps:extract-len txt))
            (if (and diam (> lenval 0.0))
              (foreach dd (ps:split-diams diam)
                (setq sys-data-map
                  (mapcar (function (lambda (sd)
                    (if (= (car sd) best-sys)
                      (list best-sys
                        (ps:accum (strcat "Pipe d" dd " m") lenval (cadr sd)))
                      sd)))
                    sys-data-map))))))))
    (setq i (1+ i)))

  ;; ---- CSV ----
  (setq results sys-data-map)
  (foreach sys results
    (princ (strcat "\n  " (car sys) ": "
                   (itoa (length (cadr sys))) " items")))

  (setq all-keys (ps:sort-keys (ps:all-keys results)))
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
(princ "\nPIPES_SYSTEM v4.0 loaded. Command: PIPES_SYSTEM")
(princ)
