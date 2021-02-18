;;; FLATTEN.LSP version 2k.0, 25-May-1999
;;;
;;; 邱枫汉化并修正
;;; 2005年9月5日
;;;
;;; FLATTEN sets the Z-coordinates of these types of objects to 0
;;; in the World Coordinate System:
;;; "3DFACE" "ARC" "ATTDEF" "CIRCLE" "DIMENSION" 
;;; "ELLIPSE" "HATCH" "INSERT" "LINE" "LWPOLYLINE"
;;; "MTEXT" "POINT" "POLYLINE" "SOLID" "TEXT"
;;;
;;;-----------------------------------------------------------------------
;;; copyright 1990-1999 by Mark Middlebrook
;;; Daedalus Consulting
;;; e-mail: markmiddlebrook@compuserve.com
;;;
;;; Thanks to Vladimir Livshiz for improvements in polyline handling
;;; and the addition of several other object types.
;;;
;;; You are free to distribute FLATTEN.LSP to others so long as you do not
;;; charge for it.
;;;
;;;-----------------------------------------------------------------------
;;;*Why Use FLATTEN?
;;;
;;; FLATTENing is useful in at least two situations:
;;; 1) You receive a DXF file created by another CAD program and discover
;;; that all the Z coordinates contain small round-off errors. These
;;; round-off errors can prevent you from object snapping to
;;; intersections and make your life difficult in other ways as well.
;;; 2) In a supposedly 2D drawing, you accidentally create one object with
;;; a Z elevation and end up with a drawing containing objects partly
;;; in and partly outside the Z=0 X-Y plane. As with the round-off
;;; problem, this situation can make object snaps and other procedures
;;; difficult.
;;;
;;; Warning: FLATTEN is not for flattening the custom objects created by
;;; applications such as Autodesk's Architectural Desktop. ADT and similar
;;; programs create "application-defined objects" that only the
;;; application really knows what to do with. FLATTEN has no idea how
;;; to handle application-defined objects, so it leaves them alone.
;;;
;;;-----------------------------------------------------------------------
;;;*How to Use FLATTEN
;;;
;;; This version of FLATTEN works with AutoCAD R12 through 2000.
;;;
;;; To run FLATTEN, load it using AutoCAD's APPLOAD command, or type:
;;; (load "FLATTEN")
;;; at the AutoCAD command prompt. Once you've loaded FLATTEN.LSP, type:
;;; FLATTEN
;;; to run it. FLATTEN will tell you what it's about to do and ask you
;;; to confirm that you really want to flatten objects in the current
;;; drawing. If you choose to proceed, FLATTEN prompts you to select objects
;;; to be flattened (press ENTER to flatten all objects in the drawing).
;;; After you've selected objects and pressed ENTER, FLATTEN goes to work.
;;; It reports the number of objects it flattens and the number left
;;; unflattenened (because they were objects not recognized by FLATTEN; see 
;;; the list of supported objects above).
;;;
;;; If you don't like the results, just type U to undo FLATTEN's work.
;;;
;;;-----------------------------------------------------------------------
;;;*Known limitations
;;; 1) FLATTEN doesn't support all of AutoCAD's object types. See above
;;; for a list of the object types that it does work on.
;;; 2) FLATTEN doesn't flatten objects nested inside of blocks.
;;; (You can explode blocks before flattening. Alternatively, you can
;;; WBLOCK block definitions to separate DWG files, run FLATTEN in
;;; each of them, and then use INSERT in the parent drawing to update
;;; the block definitions. Neither of these methods will flatten
;;; existing attributes, though.
;;; 3) FLATTEN flattens objects onto the Z=0 X-Y plane in AutoCAD's
;;; World Coordinate System (WCS). It doesn't currently support
;;; flattening onto other UCS planes.
;;;
;;;=======================================================================

(defun C:FLT (/ olderr oldcmd zeroz ss1 ss1len i
numchg numnot numno0 ssno0 ename elist etype
yorn vrt crz
)
;;Error handler
(setq olderr *error*)
(defun *error* (msg)
(if (= msg "quit / exit abort")
(princ)
(princ (strcat "error: " msg))
)
(setq *error* olderr)
(command "._UCS" "_Restore" "$FLATTEN-TEMP$"
"._UCS" "_Delete" "$FLATTEN-TEMP$"
)
(command "._UNDO" "_End")
(setvar "CMDECHO" oldcmd)
(princ)
)

;;Function to change Z coordinate to 0

(defun zeroz (key zelist / oplist nplist)
(setq oplist (assoc key zelist)
nplist (reverse (append '(0.0) (cdr (reverse oplist))))
zelist (subst nplist oplist zelist)
)
(entmod zelist)
)

;;Setup
(setq oldcmd (getvar "CMDECHO"))
(setvar "CMDECHO" 0)
(command "._UNDO" "_Group")
(command "._UCS" "_Delete" "$FLATTEN-TEMP$"
"._UCS" "_Save" "$FLATTEN-TEMP$"
"._UCS" "World"
) ;set World UCS

;;Get input
(prompt
(strcat
"\n此程序校正二维图形的Z坐标不为0的情况."
)
)

(initget "Yes No")
(setq yorn (getkword "\n是否继续: "))
(cond
((/= yorn "No")
(graphscr)
(prompt "\n选择要拍平的物体: ")
(prompt
"<直接回车将会自动选择所有物体>"
)
(setq ss1 (ssget))
(if (null ss1) ;if enter...
(setq ss1 (ssget "X")) ;select all entities in database
)

;;*initialize variables
(setq ss1len (sslength ss1) ;length of selection set
i 0 ;loop counter
numchg 0 ;number changed counter
numnot 0 ;number not changed counter
numno0 0 ;number not changed and Z /= 0 counter
ssno0 (ssadd) ;selection set of unchanged entities
) ;setq

;; 首先设置所有Thickness为0; By邱枫
(princ "\n设置Thickness为0...")
(command "_.chprop" ss1 "" "_Thickness" 0 "")

;;*do the work
(prompt "\n处理Z轴坐标中...")
(while (< i ss1len) ;while more members in the SS
(if (= 0 (rem i 10))
(princ (strcat "\r处理Z轴坐标中..." (itoa i)))
)
(setq ename (ssname ss1 i) ;entity name
elist (entget ename) ;entity data list
etype (cdr (assoc 0 elist)) ;entity type
)

;;*Keep track of entities not flattened
(if (not (member etype
'("3DFACE" "ARC" "ATTDEF"
"CIRCLE" "DIMENSION" "ELLIPSE"
"HATCH" "INSERT" "LINE"
"LWPOLYLINE" "MTEXT" "POINT"
"POLYLINE" "SOLID" "TEXT"
)
)
)
(progn ;leave others alone
(setq numnot (1+ numnot))
(if (/= 0.0 (car (reverse (assoc 10 elist))))
(progn ;add it to special list if Z /= 0
(setq numno0 (1+ numno0))
(ssadd ename ssno0)
)
)
)
)

;;Change group 10 Z coordinate to 0 for listed entity types.
(if (member etype
'("3DFACE" "ARC" "ATTDEF" "CIRCLE"
"DIMENSION" "ELLIPSE" "HATCH" "INSERT"
"LINE" "MTEXT" "POINT" "POLYLINE"
"SOLID" "TEXT"
)
)
(setq elist (zeroz 10 elist) ;change entities in list above
numchg (1+ numchg)
)
)

;;Change group 11 Z coordinate to 0 for listed entity types.
(if
(member etype
'("3DFACE" "ATTDEF" "DIMENSION" "LINE" "TEXT" "SOLID")
)
(setq elist (zeroz 11 elist))
)

;;Change groups 12 and 13 Z coordinate to 0 for SOLIDs and 3DFACEs.
(if (member etype '("3DFACE" "SOLID"))
(progn
(setq elist (zeroz 12 elist))
(setq elist (zeroz 13 elist))
)
)

;;Change groups 13, 14, 15, and 16
;;Z coordinate to 0 for DIMENSIONs.
(if (member etype '("DIMENSION"))
(progn
(setq elist (zeroz 13 elist))
(setq elist (zeroz 14 elist))
(setq elist (zeroz 15 elist))
(setq elist (zeroz 16 elist))
)
)

;;Change each polyline vertex Z coordinate to 0.
;;Code provided by Vladimir Livshiz, 09-Oct-1998
(if (= etype "POLYLINE")
(progn
(setq vrt ename)
(while (not (equal (cdr (assoc 0 (entget vrt))) "SEQEND"))
(setq elist (entget (entnext vrt)))
(setq crz (cadddr (assoc 10 elist)))
(if (/= crz 0)
(progn
(zeroz 10 elist)
(entupd ename)
)
)
(setq vrt (cdr (assoc -1 elist)))
)
)
)

;;Special handling for LWPOLYLINEs
(if (member etype '("LWPOLYLINE"))
(progn
(setq elist (subst (cons 38 0.0) (assoc 38 elist) elist)
numchg (1+ numchg)
)
(entmod elist)
)
)

(setq i (1+ i)) ;next entity
)
(prompt " 完成.")

;;Print results
(prompt (strcat "\n" (itoa numchg) " 个物体已拍平."))
(prompt
(strcat "\n" (itoa numnot) " 个物体没有拍平.")
)

;;If there any entities in ssno0, show them
(if (/= 0 numno0)
(progn
(prompt (strcat " ["
(itoa numno0)
" 个物体具有非0基点无法处理]"
)
)
(getstring
"\n回车可以亮显这些物体... "
)
(command "._SELECT" ssno0)
(getstring "\n请按回车取消亮显... ")
(command "")
)
)
)
)

(command "._UCS" "_Restore" "$FLATTEN-TEMP$"
"._UCS" "_Delete" "$FLATTEN-TEMP$"
)
(command "._UNDO" "_End")
(setvar "CMDECHO" oldcmd)
(setq *error* olderr)
(princ)
)

(prompt
"\n二维图形校正程序已加载."
)
(princ)

;;;eof 