;;; corregir_registro.lsp  -- SIN PROBAR (generado sin AutoCAD). Probar primero con copias.
;;; Reemplaza "1006-2024-2847508" por "7402-R-16-28686" en TEXT/MTEXT/ATTRIB
;;; (espacio modelo, papel y definiciones de bloque).
;;;
;;; Uso:
;;;   (load "corregir_registro.lsp")
;;;   CREG        -> dibujo actual
;;;   CREGLOTE    -> todos los DWG de una carpeta (ObjectDBX, no abre ventanas)

(vl-load-com)

(setq *creg-viejo* "1006-2024-2847508"
      *creg-nuevo* "7402-R-16-28686")

(defun creg-reemplazar (s / p)
  (while (setq p (vl-string-search *creg-viejo* s))
    (setq s (strcat (substr s 1 p) *creg-nuevo* (substr s (+ p 1 (strlen *creg-viejo*))))))
  s)

(defun creg-obj (o / n txt)
  (if (vlax-property-available-p o 'TextString)
    (progn
      (setq txt (vla-get-TextString o))
      (if (vl-string-search *creg-viejo* txt)
        (progn (vla-put-TextString o (creg-reemplazar txt)) 1)
        0))
    0))

(defun creg-coleccion (col / n)
  (setq n 0)
  (vlax-for o col
    (if (member (vla-get-ObjectName o) '("AcDbText" "AcDbMText"))
      (setq n (+ n (creg-obj o)))
      (if (and (= (vla-get-ObjectName o) "AcDbBlockReference")
               (= (vla-get-HasAttributes o) :vlax-true))
        (foreach a (vlax-invoke o 'GetAttributes)
          (setq n (+ n (creg-obj a)))))))
  n)

(defun creg-doc (doc / n)
  (setq n (creg-coleccion (vla-get-ModelSpace doc)))
  (vlax-for l (vla-get-Layouts doc)
    (setq n (+ n (creg-coleccion (vla-get-Block l)))))
  (vlax-for b (vla-get-Blocks doc)
    (if (= (vla-get-IsXRef b) :vlax-false)
      (setq n (+ n (creg-coleccion b)))))
  n)

(defun c:CREG ()
  (princ (strcat "\nTextos corregidos: " (itoa (creg-doc (vla-get-ActiveDocument (vlax-get-acad-object))))))
  (princ))

(defun c:CREGLOTE (/ carpeta archivos dbx doc n)
  (setq carpeta (getstring T "\nCarpeta con DWG (ej. C:/proyecto/planos): "))
  (setq carpeta (vl-string-right-trim "/\\" (vl-string-translate "\\" "/" carpeta)))
  (setq archivos (vl-directory-files carpeta "*.dwg" 1))
  (setq dbx (vlax-create-object
              (strcat "ObjectDBX.AxDbDocument." (substr (getvar "ACADVER") 1 2))))
  (foreach f archivos
    (setq doc (strcat carpeta "/" f))
    (vla-open dbx doc)
    (setq n (creg-doc dbx))
    (if (> n 0) (vla-saveas dbx doc))
    (princ (strcat "\n" f ": " (itoa n) " textos corregidos")))
  (vlax-release-object dbx)
  (princ))
