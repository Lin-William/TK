(princ "Author: William Lin - To Run Type : TK")

(defun c:tk (/ 
             #myCmdEcho 
             #myOsMode
             ) 
  
  (setq #myCmdEcho (getvar "CMDECHO")
        #myOsMode (getvar "OSMODE")
  ) 
  
  (defun *error* (msg) 
    (setvar "CMDECHO" #myCmdEcho)
    (setvar "OSMODE" #myOsMode)
    (princ "error:")
    (princ msg)
    (princ)
  )
  
  
  (dcl_TK)
  (setvar "CMDECHO" #myCmdEcho)
  (setvar "OSMODE" #myOsMode)
  (prin1)
)

(defun dcl_TK(/ 
              abbr
              n_zt 
              n_pm 
              n_zd
              n_cl
              n_hd
              n_jd              
              )
  
  (setq dcl_id(load_dialog "E:\\Stupefy\\lisp\\TK\\TK.dcl"));����ǰɾ��·��
  (new_dialog "TK" dcl_id)
  (action_tile "accept" "(ok_TK)(done_dialog 1)")
  (start_dialog)
  (draw_tk)
)

(defun ok_TK()
  (setq abbr (get_tile "kabbr")
        n_zt (atoi (get_tile "kzt"))
        n_pm (atoi (get_tile "kpm"))
        n_zd (atoi (get_tile "kzd"))
        n_cl (atoi (get_tile "kcl"))
        n_hd (atoi (get_tile "khd"))
        n_jd (atoi (get_tile "kjd"))
        village (get_tile "kname")
        )
)

(defun draw_tk(            
              /
              tk 
              pnt_base 
              pnt_target
              gap_pm     ;;���
              baseTH_x   ;;ͼ�ź���
              baseTH_y   ;;ͼ������
              baseTM_X   ;;ͼ������
              baseTM_y   ;;ͼ������
              style      ;;��ʽ
              base_x
              base_y
              target_x
              target_y
              obj_lst
              $tk
              i
              num
              x
              y
              text
              num_pic
              )
  ; ��ʼ������
  (setq gap_pm   20       
        baseTH_x 537.93   
        baseTH_y 13.79    
        baseTM_x 538.04  
        baseTM_y 28.94   
        style (itoa 4)
        
        tk (ssget)
        pnt_base (getpoint "\nָһ��ͼ�����½�: ")
        pnt_target (getpoint "\n���Ķ�: ")
        
        base_x   (nth 0 pnt_base)
        base_y   (nth 1 pnt_base)
        target_x (nth 0 pnt_target)
        target_y (nth 1 pnt_target)
        i 1
  )
  
  (setvar "CMDECHO" 0)
  (setvar "OSMODE" 0)
  
  ;��Ǹ���ǰ���һ��ʵ��
  (setq obj_lst (entlast))
  (command "copy" tk "" (list base_x base_y) (list target_x target_y) "")
  ;;���β��ұ�Ǻ�ʵ�������ѡ��
  (setq $tk (ssadd))
  (while (setq obj_lst (entnext obj_lst)) 
    (ssadd obj_lst $tk)
  )
  (command "copy" 
           $tk
           ""
           (list target_x target_y)
           "A"
           n_pm
           (list (+ 594 gap_pm target_x) target_y)
           ""
  )
  
  (while (<= i n_pm) 
    (setq num  (itoa i)
          x    (+ 260 target_x (* (+ 594 gap_pm) (- i 1)))
          y    (+ 50 target_y)
          text (strcat village "��ˮ����ƽ��ͼ(" num "/" (itoa n_pm) ")")
    )

    (entmake 
      (list '(0 . "MTEXT") 
            '(100 . "AcDbEntity")
            '(100 . "AcDbMText")
            '(8 . "JustOneLastTime")
            (cons 7 style)
            (cons 1 text)
            (cons 10 (list x y 0))
            (cons 40 7.5)
      )
    )

    (setq x (+ baseTM_x target_x (* (+ 594 gap_pm) (- i 1)))
          y (+ baseTM_y target_y)
    )

    (entmake 
      (list '(0 . "MTEXT") 
            '(100 . "AcDbEntity")
            '(100 . "AcDbMText")
            '(8 . "JustOneLastTime")
            (cons 7 style)
            (cons 1 text)
            (cons 10 (list x y 0))
            (cons 40 4)
            (cons 71 5)
      )
    )

    (if (< i 9) 
      (setq num_pic (strcat "0" (itoa (+ i 1))))
      (setq num_pic (itoa (+ i 1)))
    )

    (setq x    (+ baseTH_x target_x (* (+ 594 gap_pm) (- i 1)))
          y    (+ baseTH_y target_y)
          text (strcat abbr "-" num_pic)
    )

    (entmake 
      (list '(0 . "MTEXT") 
            '(100 . "AcDbEntity")
            '(100 . "AcDbMText")
            (cons 7 style)
            '(8 . "JustOneLastTime")
            (cons 1 text)
            (cons 10 (list x y 0))
            (cons 40 4)
            (cons 71 5)
      )
    )
    (setq i (+ i 1))
  )
  
)
