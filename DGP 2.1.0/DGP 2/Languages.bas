Attribute VB_Name = "Languages"
Public Function language(lang As Byte)
Select Case lang
    Case 1
    main.mnufile.Caption = "&File"
    main.mnusaveas.Caption = "&Save as..."
    main.mnuprint.Caption = "Print"
    main.mnuexit.Caption = "&Exit"
    main.mnuoptions.Caption = "&Options"
    main.mnugrid.Caption = "Grid"
    main.mnugridon.Caption = "On"
    main.mnugridoff.Caption = "Off"
    main.mnuaddlabel.Caption = "A&dd label"
    main.mnucalculator.Caption = "&Calculator"
    main.mnucolor.Caption = "C&olor..."
    main.mnuerase.Caption = "&Eraser"
    main.mnugraph.Caption = "G&paper papoer..."
    main.mnutrace.Caption = "&Trace..."
    main.mnuquality.Caption = "&Quality..."
    main.mnudetailed.Caption = "Detailed history..."
    main.mnusetting.Caption = "&Settings"
    main.mnuautocorrection.Caption = "Autocorrection"
    main.mnuautoon.Caption = "On"
    main.mnuautooff.Caption = "Off"
    main.mnuaction.Caption = "Actio&n"
    main.mnuplot.Caption = "&Plot"
    main.mnuclear.Caption = "Clear all"
    main.mnuscales.Caption = "&Scales"
    main.mnuextend.Caption = "Extend"
    main.mnuey.Caption = "in Y axis"
    main.mnuex.Caption = "in X axis"
    main.mnusqueez.Caption = "Squeeze"
    main.mnusy.Caption = "in Y axis"
    main.mnusx.Caption = "in X axis"
    main.mnuzoombox.Caption = "Zoom box"
    main.mnuhelp.Caption = "&Help"
    main.mnuhelp1.Caption = "Contents..."
    main.mnucontact.Caption = "Contact..."
    main.mnuabout.Caption = "About..."
    main.cmdplot.Caption = "PLOT"
    main.cmdPrint.Caption = "PRINT"
    main.cmdpaper.Caption = "Graph paper"
    main.cmdclrall.Caption = "Clear all"
    main.chklabel.Caption = "ADD LABEL"
    main.chkeraser.Caption = "ERASER"
    main.chktrace.Caption = "TRACE"
    main.cmddetailed.Caption = "detailed"
    main.Label2.Caption = "X from                   to                      Y from                  to"
    main.Label3.Caption = "X from                          to"
    main.Label4.Caption = "Y from                          to"
    main.Label1.Caption = "Number of ticks fo x axis                for y axis "
    main.cmdcalc.Caption = "CALCULATOR"
    main.cmdcolor.Caption = "COLOR"
    main.cmdquality.Caption = "QUALITY"
    main.cmdrefresh.Caption = "REFRESH"
    calculator.optdeg.Caption = "Degrees"
    calculator.optrad.Caption = "Radians"
    calculator.cmdeval.Caption = "Evaluate"
    calculator.cmdcancel.Caption = "Cancel"
    calculator.cmdplace.Caption = "Place"
    calculator.Caption = "Calculator"
    frmOptions.Option1(4).Caption = "Excellent quality but very slow"
    frmOptions.Option1(3).Caption = "Good quality"
    frmOptions.Option1(2).Caption = "Average quality,fast"
    frmOptions.Option1(1).Caption = "Poor quality only suitalbe if using lines"
    frmOptions.Label1.Caption = "Use lines"
    frmOptions.Label2.Caption = "Use Dots"
    frmOptions.cmdapply.Caption = "APPLY"
    frmOptions.cmdcancel.Caption = "CANCEL"
    frmOptions.Frame1.Caption = "Accuracy"
    frmOptions.Frame2.Caption = "Lines/Dots"
    frmOptions.Caption = "Detail level"
    gpaper1.Caption = "Create a graph paper"
    gpaper1.fraNumLines = "Number of lines"
    gpaper1.fraThickLines = "Lines per square"
    gpaper1.fraLines.Caption = "Design"
    gpaper1.fraScale.Caption = "Scale"
    gpaper1.fraPrint.Caption = "Printer properties"
    gpaper1.lblLinesH.Caption = "Horisontal"
    gpaper1.lblLinesV.Caption = "Vertical"
    gpaper1.lblThickV.Caption = "Vertical"
    gpaper1.lblThivkH.Caption = "Horisontal"
    gpaper1.lblThickness.Caption = "Thikness"
    gpaper1.lblColor.Caption = "Color"
    gpaper1.optInch.Caption = "Inches"
    gpaper1.optCentimeter.Caption = "Centimeters"
    gpaper1.lblcopies.Caption = "Number of copies"
    gpaper1.cmdPrint.Caption = "Print"
    gpaper1.cmdOK.Caption = "OK"
    historydet.cmdcomb.Caption = "COMBINE"
    trace.Frame1.Caption = "Choose equation to trace"
    Case 2
    main.mnufile.Caption = "&Bestand"
    main.mnusaveas.Caption = "&Opslaan als..."
    main.mnuprint.Caption = "Afdrukken"
    main.mnuexit.Caption = "&Afsluiten"
    main.mnuoptions.Caption = "&Opties"
    main.mnugrid.Caption = "Rooster"
    main.mnugridon.Caption = "Aan"
    main.mnugridoff.Caption = "Uit"
    main.mnuaddlabel.Caption = "Etiket toevoegen"
    main.mnucalculator.Caption = "&Rekenmachine"
    main.mnucolor.Caption = "Kleur..."
    main.mnuerase.Caption = "&Wisser"
    main.mnugraph.Caption = "Grafiek..."
    main.mnutrace.Caption = "&Traceren..."
    main.mnuquality.Caption = "&Kwaliteit..."
    main.mnudetailed.Caption = "Gedetailleerde geschiedenis..."
    main.mnusetting.Caption = "&Instellingen"
    main.mnuautocorrection.Caption = "Auto-correctie"
    main.mnuautoon.Caption = "Aan"
    main.mnuautooff.Caption = "Uit"
    main.mnuaction.Caption = "Handeling"
    main.mnuplot.Caption = "&Uittekenen"
    main.mnuclear.Caption = "Alles wissen"
    main.mnuscales.Caption = "&Verhoudingen"
    main.mnuextend.Caption = "Uitstrekken"
    main.mnuey.Caption = "In Y-as"
    main.mnuex.Caption = "In X-as"
    main.mnusqueez.Caption = "Inkrimpen"
    main.mnusy.Caption = "In Y-as"
    main.mnusx.Caption = "In X-as"
    main.mnuzoombox.Caption = "Zoom venster"
    main.mnuhelp.Caption = "&Help"
    main.mnuhelp1.Caption = "Inhoud..."
    main.mnucontact.Caption = "Contacteren..."
    main.mnuabout.Caption = "Over..."
    main.cmdplot.Caption = "UITTEKENEN"
    main.cmdPrint.Caption = "AFDRUKKEN"
    main.cmdpaper.Caption = "Grafiek"
    main.cmdclrall.Caption = "Alles wissen"
    main.chklabel.Caption = "TOEVOEGEN"
    main.chkeraser.Caption = "WISSER"
    main.chktrace.Caption = "TRACEREN"
    main.cmddetailed.Caption = "Gedetailleerde geschiedenis"
    main.Label2.Caption = "X van                     tot                      Y van                   tot"
    main.Label3.Caption = "X van                          tot"
    main.Label4.Caption = "Y van                          tot"
    main.Label1.Caption = "Number of ticks for x axis                for y axis "
    main.Label1.Caption = "Aantal ticks    In  x -as                     In y - as "
    main.cmdcalc.Caption = "REKENMACHINE"
    main.cmdcolor.Caption = "KLEUR"
    main.cmdquality.Caption = "KWALITEIT"
    calculator.optdeg.Caption = "Graden"
    calculator.optrad.Caption = "Radialen"
    calculator.cmdeval.Caption = "Evalueren"
    calculator.cmdcancel.Caption = "Annuleren"
    calculator.cmdplace.Caption = "Plaatsen"
    calculator.Caption = "Rekenmachine"
    main.cmdrefresh.Caption = "Vernieuwen"
    frmOptions.Option1(4).Caption = "Uitstekende kwaliteit, maar traag"
    frmOptions.Option1(3).Caption = "Goede kwaliteit"
    frmOptions.Option1(2).Caption = "Gewone kwaliteit, snel"
    frmOptions.Option1(1).Caption = "Lage kwaliteit, enkel bruikbaar bij lijnen"
    frmOptions.Label1.Caption = "Gebruik lijnen"
    frmOptions.Label2.Caption = "Gebruik punten"
    frmOptions.cmdapply.Caption = "UITVOEREN"
    frmOptions.cmdcancel.Caption = "Annuleren"
    frmOptions.Frame1.Caption = "Kwaliteit"
    frmOptions.Frame2.Caption = "Lijnen/Punten"
    frmOptions.Caption = "Nauwkeurigheid"
    gpaper1.Caption = "Grafiek"
    gpaper1.fraNumLines = "Aantal lijnen"
    gpaper1.fraThickLines = "Lijnen per vierkant"
    gpaper1.fraLines.Caption = "Uitzicht van de lijnen"
    gpaper1.fraScale.Caption = "Schaal"
    gpaper1.fraPrint.Caption = "Printer instellingen"
    gpaper1.lblLinesH.Caption = "Horizontal"
    gpaper1.lblLinesV.Caption = "Vertikaal"
    gpaper1.lblThickV.Caption = "Vertikaal"
    gpaper1.lblThivkH.Caption = "Horizontal"
    gpaper1.lblThickness.Caption = "Dikte"
    gpaper1.lblColor.Caption = "Kleur"
    gpaper1.optInch.Caption = "Inches"
    gpaper1.optCentimeter.Caption = "Centimeters"
    gpaper1.lblcopies.Caption = "Aantal kopieën"
    gpaper1.cmdPrint.Caption = "Afdrukken"
    gpaper1.cmdOK.Caption = "OK"
    historydet.cmdcomb.Caption = "SAMENVOEGEN"
    trace.Frame1.Caption = "Kies de vergelijking om op te volgen"
    Case 3
    main.mnufile.Caption = "&Fichier"
    main.mnusaveas.Caption = "&Sauver comme..."
    main.mnuprint.Caption = "Imprimer"
    main.mnuexit.Caption = "&Fermer"
    main.mnuoptions.Caption = "&Options"
    main.mnugrid.Caption = "Quadrillage"
    main.mnugridon.Caption = "Activé"
    main.mnugridoff.Caption = "Déactivé"
    main.mnuaddlabel.Caption = "Ajouter étiquette"
    main.mnucalculator.Caption = "&Calculatrice"
    main.mnucolor.Caption = "Couleur..."
    main.mnuerase.Caption = "&Torchon"
    main.mnugraph.Caption = "Graphique..."
    main.mnutrace.Caption = "&Tracer..."
    main.mnuquality.Caption = "&Qualité..."
    main.mnudetailed.Caption = "Histoire détaillée..."
    main.mnusetting.Caption = "&Attributs"
    main.mnuautocorrection.Caption = "Auto correction"
    main.mnuautoon.Caption = "Activé"
    main.mnuautooff.Caption = "Déactivé"
    main.mnuaction.Caption = "Actio&n"
    main.mnuplot.Caption = "&Calculer"
    main.mnuclear.Caption = "Essuyer tout"
    main.mnuscales.Caption = "&Echelles"
    main.mnuextend.Caption = "Etendre"
    main.mnuey.Caption = "En Y-as"
    main.mnuex.Caption = "En X-as"
    main.mnusqueez.Caption = "Réduire"
    main.mnusy.Caption = "En Y-as"
    main.mnusx.Caption = "En X-as"
    main.mnuzoombox.Caption = "Fenêtre d'agrandissement"
    main.mnuhelp.Caption = "&Aide"
    main.mnuhelp1.Caption = "Contenu..."
    main.mnucontact.Caption = "Contacter..."
    main.mnuabout.Caption = "DGP info..."
    main.cmdplot.Caption = "CALCULER"
    main.cmdPrint.Caption = "IMPRIMER"
    main.cmdpaper.Caption = "Graphique"
    main.cmdclrall.Caption = "Essuyer tout"
    main.chklabel.Caption = "ÉTIQUETTE"
    main.chkeraser.Caption = "TORCHON"
    main.chktrace.Caption = "TRACER"
    main.cmddetailed.Caption = "Histoire détaillée"
    main.Label2.Caption = "X De                      À                       Y De                     À"
    main.Label3.Caption = "X De                             À"
    main.Label4.Caption = "Y De                             À"
    main.Label1.Caption = "Nombre de ticks En X -as                  En y -as "
    main.cmdcalc.Caption = "CALCULATRICE"
    main.cmdcolor.Caption = "COULEUR"
    main.cmdquality.Caption = "QUALITÉ"
    main.cmdrefresh.Caption = "REFRAÎCHIR"
    calculator.optdeg.Caption = "Degrées"
    calculator.optrad.Caption = "Radials"
    calculator.cmdeval.Caption = "Évaluent"
    calculator.cmdcancel.Caption = "Annuler"
    calculator.cmdplace.Caption = "Placer"
    calculator.Caption = "Calculatrice"
    frmOptions.Option1(4).Caption = "Qualité excellente, mais lente"
    frmOptions.Option1(3).Caption = "Bonne qualité"
    frmOptions.Option1(2).Caption = "Qualité normale"
    frmOptions.Option1(1).Caption = "Qualité inferieur, seulement à utilser avec des lignes"
    frmOptions.Label1.Caption = "Use des lignes"
    frmOptions.Label2.Caption = "Utliser des points"
    frmOptions.cmdapply.Caption = "APPLIQUER"
    frmOptions.cmdcancel.Caption = "ANNULER"
    frmOptions.Frame1.Caption = "PRÉCISION"
    frmOptions.Frame2.Caption = "Points/Lignes"
    frmOptions.Caption = "Précision"
    gpaper1.Caption = "Graphique"
    gpaper1.fraNumLines = "Nombre de lignes"
    gpaper1.fraThickLines = "Lignes par carré"
    gpaper1.fraLines.Caption = "Perspective des lignes"
    gpaper1.fraScale.Caption = "Echelle"
    gpaper1.fraPrint.Caption = "Attributs d'imprimante"
    gpaper1.lblLinesH.Caption = "Horizontal"
    gpaper1.lblLinesV.Caption = "Vertical"
    gpaper1.lblThickV.Caption = "Vertical"
    gpaper1.lblThivkH.Caption = "Horizontal"
    gpaper1.lblThickness.Caption = "Epaisseur"
    gpaper1.lblColor.Caption = "Couleur"
    gpaper1.optInch.Caption = "Inches"
    gpaper1.optCentimeter.Caption = "Centimeters"
    gpaper1.lblcopies.Caption = "Nombre de copies"
    gpaper1.cmdPrint.Caption = "Imprimer"
    gpaper1.cmdOK.Caption = "OK"
    historydet.cmdcomb.Caption = "COMBINER"
    trace.Frame1.Caption = "Selectionnez l'équation à tracer."
End Select
End Function
