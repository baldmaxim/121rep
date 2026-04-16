;;; =============================================================
;;; PIPES_SYSTEM.lsp  v2.0
;;; HVAC refrigerant system analyzer for AutoCAD
;;;
;;; Finds: ODU (ARUN*), IDU (IAC*/ARNU*), Refnets (ARBLN*),
;;;        pipe lengths by diameter
;;; Groups by system name from MTEXT near ODU block.
;;; Traces connectivity via BFS (LINE/LWPOLYLINE/POLYLINE + blocks).
;;; Exports to CSV (semicolon-separated, comma decimals for RU Excel).
;;;
;;; Command: PIPES_SYSTEM
;;; =============================================================

(vl-load-com)

;;; ---- Settings ----
(setq *PS:TOL*   3.0)    ; connection tolerance (drawing units)
(setq *PS:MRAD* 600.0)   ; MTEXT search radius near equipment block
(setq *PS:PRAD* 200.0)   ; pipe label search radius

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

;;; starts-with (case-insensitive)
(defun ps:sw (s pfx)
  (and (>= (strlen s) (strlen pfx))
       (= (strcase (substr s 1 (strlen pfx))) (strcase pfx))))

;;; Parse MTEXT raw string: strip RTF codes, split by \P -> list of lines
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
;;;  AutoCAD entity helpers
;;; ================================================================

(defun ps:etype (e) (cdr (assoc 0  (entget e))))
(defun ps:bname (e) (cdr (assoc 2  (entget e))))
(defun ps:ipt   (e) (cdr (assoc 10 (entget e))))

;;; All 2D vertices of LINE / LWPOLYLINE / POLYLINE / INSERT
(defun ps:verts (e / ed et pts sub)
  (setq ed (entget e)  et (cdr (assoc 0 ed)))
  (cond
    ((= et "LINE")
     (list (cdr (assoc 10 ed)) (cdr (assoc 11 ed))))
    ((= et "LWPOLYLINE")
     (setq pts nil)
     (foreach pair ed
       (if (= (car pair) 10) (setq pts (append pts (list (cdr pair))))))
     pts)
    ((= et "POLYLINE")
     (setq pts nil  sub (entnext e))
     (while (and sub (not (= (ps:etype sub) "SEQEND")))
       (if (= (ps:etype sub) "VERTEX")
         (setq pts (append pts (list (cdr (assoc 10 (entget sub)))))))
       (setq sub (entnext sub)))
     pts)
    ((= et "INSERT") (list (cdr (assoc 10 ed))))
    (t nil)))

;;; Centroid of a point list
(defun ps:center (pts / n sx sy)
  (if (null pts) nil
    (progn
      (setq n (length pts)  sx 0.0  sy 0.0)
      (foreach p pts (setq sx (+ sx (car p))  sy (+ sy (cadr p))))
      (list (/ sx n) (/ sy n) 0.0))))

;;; 2D distance
(defun ps:d2 (a b)
  (sqrt (+ (expt (- (car  a) (car  b)) 2)
           (expt (- (cadr a) (cadr b)) 2))))

;;; Any point pair within tolerance?
(defun ps:touch? (v1 v2 tol / found)
  (setq found nil)
  (foreach a v1
    (foreach b v2
      (if (and a b (<= (ps:d2 a b) tol)) (setq found T))))
  found)

;;; ================================================================
;;;  MTEXT search and parsing
;;; ================================================================

;;; Find nearest MTEXT within radius containing substring key
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

;;; Parse equipment MTEXT -> alist ((key . value) ...)
;;; Line format: "Key : Value"
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

;;; Find value in alist by partial key match (case-insensitive ASCII fragment)
(defun ps:find-val (frag alist / found)
  (setq found nil)
  (foreach p alist
    (if (and (null found)
             (vl-string-search (strcase frag) (strcase (car p))))
      (setq found (cdr p))))
  found)

;;; Find first value that starts with prefix (e.g. "ARUN", "ARNU")
(defun ps:find-val-sw (prefix alist / found)
  (setq found nil)
  (foreach p alist
    (if (and (null found) (ps:sw (cdr p) prefix))
      (setq found (cdr p))))
  found)

;;; ================================================================
;;;  Block attributes and IDU model detection
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

;;; IDU model: 1) attribute starting ARNU  2) nearby MTEXT  3) block name
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
                    (setq model (if sp (substr sub 1 sp) sub))))))))))))
  (if (and model (> (strlen model) 0)) model bname))

;;; ================================================================
;;;  Pipe label search
;;; ================================================================

;;; Diameter label: contains ":", no "/", no "ID", above center
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

;;; Length label: contains "/", below or at center Y
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
;;;  BFS connectivity tracing
;;; ================================================================

(defun ps:bfs (start-ent pool tol / queue visited result cur cv cand cv2)
  (setq queue   (list start-ent)
        visited (list start-ent)
        result  (list start-ent))
  (while queue
    (setq cur   (car queue)
          queue (cdr queue)
          cv    (ps:verts cur))
    (foreach cand pool
      (if (not (member cand visited))
        (progn
          (setq cv2 (ps:verts cand))
          (if (ps:touch? cv cv2 tol)
            (progn
              (setq visited (append visited (list cand))
                    queue   (append queue   (list cand))
                    result  (append result  (list cand)))))))))
  result)

;;; ================================================================
;;;  Result accumulation
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

;;; ================================================================
;;;  CSV export
;;; ================================================================

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

(defun C:PIPES_SYSTEM (/ ss i e ed bname ipt me parsed
                          sys-name model systems pool
                          sys sys-ents sys-data
                          etype everts cen diam lenstr lenval
                          imodel key results all-keys fname)
  (vl-load-com)
  (princ "\n=== PIPES_SYSTEM v2.0 ===")

  ;; 1. Build pool: lines + ARBLN* + IAC* blocks
  (princ "\nCollecting objects...")
  (setq pool nil)

  (setq ss (ssget "_X" '((0 . "LINE,LWPOLYLINE,POLYLINE"))))
  (if ss
    (progn (setq i 0)
      (repeat (sslength ss)
        (setq pool (append pool (list (ssname ss i))))
        (setq i (1+ i)))))

  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn (setq i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i)  bname (ps:bname e))
        (if (or (ps:sw bname "ARBLN") (ps:sw bname "IAC"))
          (setq pool (append pool (list e))))
        (setq i (1+ i)))))

  (princ (strcat " " (itoa (length pool)) " objects found."))

  ;; 2. Find ODU blocks ARUN* -> system names
  (princ "\nSearching ODU blocks (ARUN*)...")
  (setq systems nil)

  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn (setq i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i)  bname (ps:bname e))
        (if (ps:sw bname "ARUN")
          (progn
            (setq ipt (ps:ipt e))
            (setq me (ps:mtext-near ipt *PS:MRAD* "ID"))
            (if me
              (progn
                (setq parsed   (ps:parse-mtext me))
                ;; System name: value after "ID" key
                (setq sys-name (ps:find-val "ID" parsed))
                ;; Model: first value starting with "ARUN"
                (setq model    (ps:find-val-sw "ARUN" parsed)))
              (setq sys-name bname  model bname))
            (if (null sys-name) (setq sys-name bname))
            (if (null model)    (setq model    bname))
            (setq systems (append systems (list (list sys-name model e))))))
        (setq i (1+ i)))))

  (if (null systems)
    (progn
      (alert "No ARUN* blocks found! Check drawing.")
      (exit)))

  (princ (strcat " " (itoa (length systems)) " systems found."))

  ;; 3. Trace each system, collect data
  (setq results nil)

  (foreach sys systems
    (setq sys-name (car   sys)
          model    (cadr  sys)
          e        (caddr sys))

    (princ (strcat "\n  Tracing: " sys-name " ..."))
    (setq sys-ents (ps:bfs e pool *PS:TOL*))
    (princ (strcat " " (itoa (length sys-ents)) " objects."))

    (setq sys-data nil)

    ;; ODU: always 1
    (setq sys-data (ps:accum (strcat "ODU " model) 1 sys-data))

    (foreach ent sys-ents
      (setq etype (ps:etype ent))
      (cond
        ((= etype "INSERT")
         (setq bname (ps:bname ent))
         (cond
           ((ps:sw bname "ARBLN")
            (setq sys-data (ps:accum (strcat "Refnet " bname) 1 sys-data)))
           ((ps:sw bname "IAC")
            (setq imodel (ps:vb-model ent bname))
            (setq sys-data (ps:accum (strcat "IDU " imodel) 1 sys-data)))))

        ((member etype '("LINE" "LWPOLYLINE" "POLYLINE"))
         (setq everts (ps:verts ent)
               cen    (ps:center everts))
         (if cen
           (progn
             (setq diam   (ps:diam-above cen *PS:PRAD*))
             (setq lenstr (ps:len-below  cen *PS:PRAD*))
             (if (and diam lenstr)
               (progn
                 (setq lenval (ps:extract-len lenstr))
                 (if (> lenval 0.0)
                   (setq sys-data
                     (ps:accum (strcat "Pipe " diam " m") lenval sys-data))))))))))

    (setq results (append results (list (list sys-name sys-data)))))

  ;; 4. Save CSV
  (setq all-keys (ps:all-keys results))

  (setq fname (getfiled "Save CSV report" (getvar "DWGPREFIX") "csv" 1))
  (if (null fname)
    (progn (princ "\nCancelled.") (exit)))

  (if (ps:write-csv fname results all-keys)
    (progn
      (princ (strcat "\nDone! File saved: " fname))
      (startapp "explorer" fname))
    (alert "File write error!"))

  (princ "\n=== PIPES_SYSTEM done ===")
  (princ))

;;; ================================================================
(princ "\nPIPES_SYSTEM v2.0 loaded. Command: PIPES_SYSTEM")
(princ)
