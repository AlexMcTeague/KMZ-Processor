; Written by Alex McTeague
(defun c:ImportGeoCSV (/ data file line lat lon nodeType nodeName len nodeType nodeName len wcsPoint blkName)
  (setq prevOSMode (getvar "osmode"))
  (setq prevLayer (getvar "clayer"))
  (setq prevColor (getvar "cecolor"))
  
  (if
    (and
      ; Ask the user to select a CSV file, and double check that it's readable
      (setq file (getfiled "Select CSV File" "" "csv" 16))
      (setq data (LM:readcsv file))
    )
    (progn
      (princ "\n")
      ; Loop through the lines in the CSV
      (foreach line (cdr data)
        (setq blkName nil)
        (setq wcsPoint nil)
        
        ; Extract each individual piece of data
        (setq lat (nth 0 line)) ; Latitude
        (setq lon (nth 1 line)) ; Longitude
        (setq nodeType (nth 2 line)) ; Node type
        (setq nodeName (nth 3 line)) ; Node name
        (setq len (nth 4 line)) ; Connection length. Returns nil if doesn't exist
        (setq fullCoords (nth 5 line)) ; String containing all coordinate key-pairs. Each pair is separated by ' ', each value is separated by ','
      
        ; Choose a block based on node type, or perform another operation
        (cond
          ; sic: two spaces in Joint Use Transformer
          ((wcmatch nodeType "CUT MARK,Comms Pole,Guy Pole,Power,Joint Use,Joint Use  Transformer,Steel Transmission {Secondary},Transmission (Secondary),Transformer Pole")
            (setq blkName "POLE_EXISTING")
          )
          ((wcmatch nodeType "PED")
            ; TODO
          )
          ((wcmatch nodeType "VAULT")
            ; TODO
          )
          ((wcmatch nodeType "Aerial Strand,overhead guy,overlash,slack span")
            (if (and fullCoords len)
              (progn
                (PlaceCable fullCoords len "Aerial")
              )
            )
          )
          ((wcmatch nodeType "underground cable")
            (if (and fullCoords len)
              (progn
                (PlaceCable fullCoords len "Underground")
              )
            )
          )
        )
        
        ; Place the chosen block at the node's coordinates
        (if blkName
          (progn
            ; Convert GPS coordinates to WCS coordinates (x y z)
            (setq wcsPoint (ConvertGPSToWCS (list lat lon) nil))
                    
            (if wcsPoint
              (progn
                ;(princ (strcat "\nLat: " lat ", Long: " lon ", Model coords: " (vl-princ-to-string wcsPoint) "\n"))
                (progn
                  ; Turn off Object Snapping, which messes with object placement
                  (setq prevOSMode (getvar "osmode"))
                  (setvar 'osmode 0)
                  
                  ; Insert the block
                  (setvar "ATTDIA" 0) ; Disable attribute dialog when placing blocks (not always necessary, but it can help)
                  (command "_INSERT" blkName wcsPoint 1 1 0) ; (1/1 Scale, 0 rotation)
                  (setvar "ATTDIA" 1)
                  
                  ; Restore Object Snapping to its previous settings
                  (setvar 'osmode prevOSMode)
                )
              )
              (alert "Coordinate conversion failed. Check drawing's coordinate system.")
            )
          )
        )
      )
    )
  )
  (princ)
)

(defun PlaceCable (fullCoords dist kind /)
  ; Break the fullCoords string into a list of coordinate pairs
  (setq coordList (LM:str->lst fullCoords " "))
  ;(princ (strcat "\n" (vl-princ-to-string coordList) "\n"))

  ; Loop through the coordinate pairs
  (setq pt-list nil)
  (setq start-pt nil)
  (setq end-pt nil)
  (foreach pairString coordList
    ; Convert the coordinate pair string into a point (list)
    (setq wcsPoint (ConvertGPSToWCS (LM:str->lst pairstring ",") T))

    ; Sets the start point (once) and last point (wherever the loop ends, end-pt will stay there)
    (if (not start-pt)
      (setq start-pt wcsPoint)
    )
    (setq end-pt wcsPoint)
    
    ; Add the point to a list of points
    (setq pt-list (cons wcsPoint pt-list)) 
  )
  
  ; Turn off Object Snapping, which messes with object placement
  (setvar 'osmode 0)
    
  ; Pass the list of points to the Polyline command to draw it
  (princ (strcat "\n" (vl-princ-to-string pt-list) "\n"))
  (command "_.PLINE")
  (foreach pt pt-list
    (command pt)
  )
  (command "") ; End command

  ; Restore Object Snapping to its previous settings
  (setvar 'osmode prevOSMode)
  
  ; Get components of the start/end points
  (setq pt1X (car start-pt))
  (setq pt1Y (cadr start-pt))
  (setq pt2X (car end-pt))
  (setq pt2Y (cadr end-pt))
  
  ; Get the label rotation so the text is always face-up
  (if (= pt2X pt1X)
    (if (> pt2Y pt1Y)
      (setq angleRad (* pi 0.5))
      (setq angleRad (* pi 1.5))
    )
    (setq angleRad (atan (/ 
      (- pt2Y pt1Y)
      (- pt2X pt1X)
    )))
  )
  (setq angleDeg (* angleRad (/ 180 pi)))
  
  ; Get midpoint
  (setq midpoint
    (list
      (/ (+ (car start-pt) (car end-pt)) 2.0)
      (/ (+ (cadr start-pt) (cadr end-pt)) 2.0)
    )
  )
  
  ; Change layer/color/snapping
  (cond
    (= kind "Aerial"
      (setq layer "Aerial Fiber")
    )
    (= kind "Underground"
      (setq layer "UNDERGROUND")
    )
  )
  (command "_.CLAYER" layer)
  (command "_.CECOLOR" "ByLayer")
  (setvar 'osmode 0)
  
  ; Create the distance label at the midpoint of the new line
  (setq textHeight 7)
  ; NOTE: This won't work properly if you have a default text height set with the ST command
  (command "_.TEXT" "J" "M" midPoint textHeight angleDeg dist)
  
  ; Return to previous state
  (setvar 'osmode prevOSMode)
  (command "_.CLAYER" prevLayer)
  (command "_.cecolor" prevColor)
  
  (princ)
)

(defun ConvertGPSToWCS (gpsCoords reversed /)
  ; NOTE: T for true, nil for false
  (if reversed
    (progn
      (setq flippedCoords (list (cadr gpsCoords) (car gpsCoords)))
      (setq gpsCoords flippedCoords)
    )
  )
  
  ; ACAD LT doesn't have a function to directly convert GPS coords to Model Space coords.
  ; Instead, we place a temporary LatLong object with GPS coords
  ; Then we get its model space coords, and delete the temp object
  (command "_GEOMARKLATLONG" (car gpsCoords) (cadr gpsCoords) "")
  (setq ent (entlast))
  (setq wcsPoint (cdr (assoc 10 (entget ent))))
  (entdel ent)
  (setq wcsPoint wcsPoint)
)

(defun LM:str->lst ( str del / pos )
    (if (setq pos (vl-string-search del str))
        (cons (substr str 1 pos) (LM:str->lst (substr str (+ pos 1 (strlen del))) del))
        (list str)
    )
)