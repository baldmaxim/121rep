;;; =============================================================
;;; SUM_PIPES.lsp
;;; Суммирование длин труб по диаметрам из текстовых меток.
;;;
;;; Ожидаемый формат пар текстов на чертеже:
;;;   Верхний текст (диаметр):  "6,35:12,7"
;;;   Нижний текст  (длина):    "0,6 / 3,2 m(0)"
;;;
;;; Команда: SUM_PIPES
;;; =============================================================

(vl-load-com)

;;; Удаление пробелов в начале и конце строки
(defun sp-trim (s / i len)
  (setq i 1  len (strlen s))
  (while (and (<= i len) (= (substr s i 1) " "))
    (setq i (1+ i)))
  (setq s (substr s i))
  (setq len (strlen s))
  (while (and (> len 0) (= (substr s len 1) " "))
    (setq s (substr s 1 (1- len)))
    (setq len (strlen s)))
  s
)

;;; Получить текстовую строку из объекта TEXT или MTEXT
(defun sp-get-str (ent / ed)
  (setq ed (entget ent))
  (cdr (assoc 1 ed))
)

;;; Получить точку вставки текстового объекта
(defun sp-get-pt (ent / ed)
  (setq ed (entget ent))
  (cdr (assoc 10 ed))
)

;;; Извлечь первое число перед "/", заменив запятую на точку
(defun sp-extract-num (s / pos sub)
  (setq pos (vl-string-search "/" s))
  (if pos
    (progn
      (setq sub (sp-trim (substr s 1 pos)))
      (setq sub (vl-string-subst "." "," sub))
      (atof sub)
    )
    nil
  )
)

;;; Найти ближайший текст с ":" строго выше базовой точки в заданном радиусе
(defun sp-find-diam (base-pt radius / ss2 j ent2 txt2 pt2 d best-d best-txt x0 y0)
  (setq x0      (car  base-pt)
        y0      (cadr base-pt)
        best-d  1e10
        best-txt nil)
  ;; Выборка кандидатов: окно от (x0-r, y0) до (x0+r, y0+r) — только выше
  (setq ss2 (ssget "_C"
              (list (- x0 radius) y0              0.0)
              (list (+ x0 radius) (+ y0 radius)   0.0)))
  (if ss2
    (progn
      (setq j 0)
      (repeat (sslength ss2)
        (setq ent2  (ssname ss2 j)
              txt2  (sp-get-str ent2)
              pt2   (sp-get-pt  ent2))
        (if (and txt2
                 pt2
                 (vl-string-search ":" txt2)       ; содержит ":"
                 (null (vl-string-search "/" txt2)) ; не является строкой длины
                 (> (cadr pt2) y0))                 ; строго выше
          (progn
            (setq d (distance base-pt pt2))
            (if (< d best-d)
              (setq best-d   d
                    best-txt (sp-trim txt2)))))
        (setq j (1+ j)))))
  best-txt
)

;;; Аккумулировать: добавить num к текущей сумме по ключу key
(defun sp-accum (key num lst / p)
  (if (setq p (assoc key lst))
    (subst (cons key (+ (cdr p) num)) p lst)
    (append lst (list (cons key num)))
  )
)

;;; ================================================================
;;;  Главная команда
;;; ================================================================
(defun C:SUM_PIPES (/ ss i ent txt num diam result)
  (vl-load-com)
  (setq result nil)

  ;; Пользователь выбирает нужные тексты вручную
  (princ "\nВыберите текстовые объекты (рамкой или по одному), затем Enter: ")
  (setq ss (ssget '((0 . "TEXT,MTEXT"))))

  (if (null ss)
    (princ "\nSUM_PIPES: объекты не выбраны, команда отменена.")
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq txt (sp-get-str ent))

        ;; Обрабатываем только тексты, содержащие "/"
        (if (and txt (vl-string-search "/" txt))
          (progn
            (setq num (sp-extract-num txt))
            (if (and num (> num 0.0))
              (progn
                ;; Ищем ближайший текст диаметра выше (радиус 150 единиц)
                (setq diam (sp-find-diam (sp-get-pt ent) 150.0))
                (if (null diam) (setq diam "Не определён"))
                (setq result (sp-accum diam num result))
              )
            )
          )
        )
        (setq i (1+ i))
      )

      ;; ---- Вывод результатов в текстовую консоль (F2) ----
      (textpage)
      (princ "\n")
      (princ "================================================")
      (princ "\n     СУММАРНЫЕ ДЛИНЫ ТРУБ ПО ДИАМЕТРАМ")
      (princ "\n================================================")
      (if (null result)
        (princ "\nДанные не найдены.\nПроверьте формат: 'X,X / Y m' и наличие текста диаметра выше.")
        (foreach item result
          (princ
            (strcat "\nДиаметр: "
                    (car item)
                    "  —  Суммарная длина: "
                    (rtos (cdr item) 2 2)
                    " м"))
        )
      )
      (princ "\n================================================\n")
    )
  )
  (princ)
)

(princ "\nSUM_PIPES загружен. Введите SUM_PIPES для запуска.")
(princ)
