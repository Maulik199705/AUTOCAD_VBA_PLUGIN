;Het blok definitief vervangen door revisie G.C.Haak 10-05-2005
(defun c:vervangen ()

  (setq la nil)
(setvar "cmdecho" 0)
(setq dwgn (getvar "dwgname"))
  
  (setq la (ssget "x" '((2 . "definitief2"))))
  (if (= la nil)(setq la (ssget "x" '((2 . "definitief")))))
  (setq rla (ssget "x" '((2 . "bl-revisie"))))
  
  (if (and (= la  nil)(= rla nil))
    (progn
    (setq f (open "c:\\acad2002\\foutrapport.txt" "a"))
    (write-line (getvar "dwgname") f)
    (close f)
    );einde progn
  ) 




  (setq lan (ssname la 0))
  (setq lanam (entget lan))
  (setq laname (cdr (assoc 10 lanam)))

(command "erase" la "");erased het blok definitief
(command "-insert" "bl-revisie" laname "" "" "");plaatst het blok revisie
(command "qsave") 
)



