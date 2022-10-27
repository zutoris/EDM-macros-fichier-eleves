option explicit ' oblige à la déclaration explicite de toutes les variables du programme

'*************************************
' Déclaration des types personnalisés
'*************************************
type GroupeT ' représente un groupe (ex : l'ensemble débutant, la classe de FM CI A, les élèves d'un prof...)
  nom as string ' nom du groupe
  typeGroupe as string ' type du groupe. Peut valoir : ENSEMBLE, FM, PROF, PARCOURS, TOUS
  numAdherents as new Collection ' collection des numéros des adhérents de ce groupe
end type

type TypeGroupeT ' représente un type de groupe
  typeGroupe as string ' type du groupe. Peut valoir : ENSEMBLE, FM, PROF, PARCOURS, TOUS
  nomsGroupes as new Collection ' tous les groupes de ce type
end type

'*************************************
' Déclaration des variables globales
'*************************************
global sTxtCString as string
private const NB_EMAILS_MAX = 19 ' nombre max d'emails copiés dans le presse-papier à la fois, afin de diminuer les risques d'être considéré comme spam
private const NB_ENSEMBLES_MAX = 101 ' nombre max d'ensembles +1
private ensemblesChoisisTab(NB_ENSEMBLES_MAX) as integer ' liste des indexes des groupes sélectionnnés
private nbEnsemblesSelectionnes as integer ' nombre de groupes sélectionnés
private oDlgModele as object
private oDlgControle as object
private numAdherentsParEnsemble as new Collection ' collection des groupes, dont la clé est le nom du groupe, et la valeur est un GroupeT, qui contient une collection des numéros des adhérents de ce groupe
private annule as boolean ' indique si l'utilisateur a sélectionné le bouton 'Annule'


'************************************************************
'    Programme principal
'************************************************************
Sub main

  dim document as object : document = ThisComponent ' le fichier Calc
  dim mainSheet as object : mainSheet = document.Sheets(0) ' 1er onglet du tableur
  dim numColEnsembles as long : numColEnsembles = mainSheet.getCellRangeByName("TITRE_ENS_CHORALE").CellAddress.Column ' numéro de la colonne des ensembles
  dim numColFM as long : numColFM = mainSheet.getCellRangeByName("TITRE_FM").CellAddress.Column ' numéro de la colonne des cours de FM
  dim numColEveil as long : numColEveil = mainSheet.getCellRangeByName("TITRE_EVEIL").CellAddress.Column ' numéro de la colonne des éveils
  dim numColProf as long : numColProf = mainSheet.getCellRangeByName("TITRE_PROF").CellAddress.Column ' numéro de la colonne des professeur
  dim numAdherent as integer : numAdherent = 0 ' nombre d'adhérents, dont ceux qui sont barrés
  dim indexTab as integer
  dim adresseCellule as string
  dim valCelluleLue as string ' valeur de la cellule venant d'être lue
  dim coursFMCel as string
  dim ensemblesDansCellTab() as string ' tableau contenant tous les noms des ensembles d'une cellule
  dim nomEnsemble as string ' nom d'un ensemble
  dim nomPrenomAdherent as string ' concaténation du nom et du prénom de l'adhérent, pour ne pas compter les doublons
  dim nomPrenomPrecedentAdherent as string : nomPrenomPrecedentAdherent = "" ' conservation du du nom et du prénom du précédent adhérent

  ' détermination de la liste des groupes
  do while mainSheet.getCellByPosition(0, numAdherent+1).type <> com.sun.star.table.CellContentType.EMPTY
    numAdherent = numAdherent + 1
    if mainSheet.getCellByPosition(0, numAdherent).charStrikeout <= 0 Then ' exclusion des cellules barrées dans la 1ère colonne
    
      ' traitement des ensembles
      valCelluleLue = mainSheet.getCellByPosition(numColEnsembles, numAdherent).string
      if valCelluleLue <> "" then 
  	    ensemblesDansCellTab = split(valCelluleLue, ";") ' pour le cas où un élève est dans plusieurs ensembles
  	    for each nomEnsemble in ensemblesDansCellTab
          ajouteAdherent(numAdherent, rtrim(nomEnsemble), numAdherentsParEnsemble, "ENSEMBLE")
        next nomEnsemble	  
      end if
      
      ' traitement du cours de FM
      valCelluleLue = mainSheet.getCellByPosition(numColFM, numAdherent).string
      if valCelluleLue <> "" then 
        ajouteAdherent(numAdherent, valCelluleLue, numAdherentsParEnsemble, "FM")
      end if
      
      ' traitement du cours de jardin/éveil/parcours
      valCelluleLue = mainSheet.getCellByPosition(numColEveil, numAdherent).string
      if valCelluleLue <> "" then 
        ajouteAdherent(numAdherent, valCelluleLue, numAdherentsParEnsemble, "EVEIL")
      end if
      
      ' traitement du professeur
      valCelluleLue = mainSheet.getCellByPosition(numColProf, numAdherent).string
      if valCelluleLue <> "" then 
        ajouteAdherent(numAdherent, valCelluleLue, numAdherentsParEnsemble, "PROF")
      end if
      
      ' traitement de tous les adhérents, avec filtrage des doublons (cas des élèves inscrits à plusieurs instruments)
      nomPrenomAdherent =  mainSheet.getCellByPosition(0, numAdherent).string & "-" & mainSheet.getCellByPosition(1, numAdherent).string
      if nomPrenomAdherent <> nomPrenomPrecedentAdherent then
        ajouteAdherent(numAdherent, "tous", numAdherentsParEnsemble, "TOUS")
        nomPrenomPrecedentAdherent = nomPrenomAdherent
      end if

    end if
  loop

  ' Affichage de la boite de dialogue de sélection des ensembles
  afficheBoiteDialogueSelection(numAdherentsParEnsemble)

  ' Traitement de la sélection
  if nbEnsemblesSelectionnes > 0 then
    dim idxEns as integer
    dim adherentsSlectionnes(nbEnsemblesSelectionnes - 1) as object ' tableau des collections des adhérents dont les ensembles ont été sélectionnés
    for idxEns = 0 to nbEnsemblesSelectionnes - 1
      adherentsSlectionnes(idxEns) = numAdherentsParEnsemble.item(ensemblesChoisisTab(idxEns)).numAdherents
    next idxEns
  
    ' adresses mail des adhérents dont les ensembles ont été sélectionnés
    dim emailsSelectionnes() as string : emailsSelectionnes = recupereEmails(adherentsSlectionnes, mainSheet)
    
    restitueEmails(emailsSelectionnes)
  endif

end sub



'***************************************************************************
' Ajoute l'adhérent à la collection. Si le groupe n'existe pas, il est créé
'***************************************************************************
Sub ajouteAdherent(numeroAdherent as integer, nomGroupe as string, byRef adherentsParGroupe as new Collection, typeDeGroupe as string)
  dim nomGroupeSansEspace as string : nomGroupeSansEspace = rtrim(nomGroupe) ' nom d'un ensemble, sans espace à la fin
  dim groupeDansColl as GroupeT
	
  if estGroupeDansCollection(adherentsParGroupe, nomGroupeSansEspace) then
    ' ce groupe est déjà présent, l'adhérent est alors ajouté
    groupeDansColl = adherentsParGroupe(nomGroupeSansEspace)
    groupeDansColl.numAdherents.add(numeroAdherent)
  else
    ' groupe non présent, il est créé
    groupeDansColl.nom = nomGroupeSansEspace
    groupeDansColl.typeGroupe = typeDeGroupe
    dim numAdherentsCol as new Collection
    numAdherentsCol.add(numeroAdherent)
    groupeDansColl.numAdherents = numAdherentsCol
    adherentsParGroupe.add(groupeDansColl, nomGroupeSansEspace)
  endif
end sub


'******************************************************************
' Indique si la collection possède la clé fournie. Retourne 'true' si oui, 'false' sinon.
'******************************************************************
Function estGroupeDansCollection(byRef pColl as object, pCle as string) as boolean
	dim element as object, existe as boolean
	on local error goto ErrHandler 'gérer l'erreur
		existe = false
		element = pColl(pCle) 'si pCle n'existe pas -> erreur
		existe = true
	ErrHandler:
		'ne rien faire
	estGroupeDansCollection = existe
End Function


'******************************************************************
' Affiche la boite de dialogue permettant de choisir les ensembles
'******************************************************************
Sub afficheBoiteDialogueSelection(byRef adherentsParEnsemble as new Collection)
  ' réarangement des données pour les regrouper par type
  dim tableauTypesGroupe as new Collection ' collection des types de groupe. Chaque élément est de type TypeGroupeT
  tableauTypesGroupe = getEnsemblesParType(adherentsParEnsemble)

  ' Création de la boite vide
  dim hauteurBoite as long : hauteurBoite = 90 + adherentsParEnsemble.count / 3 * 14
  dim largeurBoite as long : largeurBoite = 400
  dlg_Creation(60, 50, largeurBoite, hauteurBoite, "Copier les adresses mails")
   
  ' Ajout du texte de présentation
  dlg_Libelle(5, 4, 250, 14, "lbl1", "Sélectionnez les ensembles pour copier les adresses mails des élèves dans le presse-papier :")

  ' Ajout des cases à cocher des ensembles
  dim absColonne1 as long : absColonne1 =  20
  dim absColonne2 as long : absColonne2 = 150
  dim absColonne3 as long : absColonne3 = 280
  dim idxCheckBox as integer : idxCheckBox = 0
  afficheToutesCheckBoxDeType("TOUS", idxCheckBox, tableauTypesGroupe, absColonne1)
  afficheLigne(absColonne1, 18 + idxCheckBox * 14, 100, 1, "ligne1")
  idxCheckBox = idxCheckBox + 1
  afficheToutesCheckBoxDeType("ENSEMBLE", idxCheckBox, tableauTypesGroupe, absColonne1)
  afficheLigne(absColonne1, 18 + idxCheckBox * 14, 100, 1, "ligne2")
  idxCheckBox = idxCheckBox + 1
  afficheToutesCheckBoxDeType("EVEIL", idxCheckBox, tableauTypesGroupe, absColonne1)
  idxCheckBox = 0
  afficheToutesCheckBoxDeType("FM", idxCheckBox, tableauTypesGroupe, absColonne2)
  idxCheckBox = 0
  afficheToutesCheckBoxDeType("PROF", idxCheckBox, tableauTypesGroupe, absColonne3)
   
  ' Ajout des boutons OK et Annuler
  dlg_Bouton(largeurBoite/2 - 60, hauteurBoite - 20, 50, 14, "dlgValide", "OK",      "ValideChoixEnsembles")
  dlg_Bouton(largeurBoite/2 + 10, hauteurBoite - 20, 50, 14, "dlgAnnule", "Annuler", "Annule")
     
  ' Affichage de la boite de dialogue une fois construite
  dlg_Affiche()

End Sub


'******************************************************************
' Affiche la boite de dialogue indiquant que la copie dans le 
' presse-papier se déroule en plusieurs fois
'******************************************************************
Sub afficheBoiteDialogueResteEmails(byVal nbCopiesFaites as integer, byVal nbCopiesTotal as integer, byVal ndAdressesDejaCopiees as integer, byVal nbTotalAdresses as integer)
  dim nbAdressesRestantesMajorees as integer : nbAdressesRestantesMajorees = nbTotalAdresses - ndAdressesDejaCopiees - 1
  if nbAdressesRestantesMajorees > NB_EMAILS_MAX then
    nbAdressesRestantesMajorees = NB_EMAILS_MAX
  endif
  
  dlg_Creation(60, 45, 155, 45, "Copie partielle des adresses - " & nbCopiesFaites & " / " & nbCopiesTotal )
   
  ' Ajout du texte de présentation
  dlg_Libelle(5, 4, 250, 14, "lbl1", NB_EMAILS_MAX & " adresses mails ont été copiées dans le presse-papier.")
  ' Ajout des boutons Copier et Annuler
  dim finMessage as string
  if nbCopiesFaites = nbCopiesTotal then
    finMessage = " dernières adresses"
  else
    finMessage = " adresses suivantes"
  endif
  dlg_Bouton( 10, 20, 95, 14, "dlgSuite", "Copier les " & nbAdressesRestantesMajorees & finMessage, "SuiteRecupereEmails")
  dlg_Bouton(115, 20, 30, 14, "dlgAnnule", "Annuler", "Annule")
  ' Affichage de la boite de dialogue une fois construite
  dlg_Affiche()
End Sub


'******************************************************************
' regroupement des groupes en fonction des types
'******************************************************************
Function getEnsemblesParType(byRef adherentsParEnsemble as new Collection)
  dim nombreEnsembles as integer : nombreEnsembles = adherentsParEnsemble.count
  dim index as integer, groupe as GroupeT
  dim tableauTypesGroupe as new Collection ' collection des types de groupe. Chaque élément est de type TypeGroupeT
  
  ' regroupement des groupes en fonction des types
  for index = 1 to nombreEnsembles
    groupe = adherentsParEnsemble.item(index)
    if estGroupeDansCollection(tableauTypesGroupe, groupe.typeGroupe) then
      tableauTypesGroupe(groupe.typeGroupe).nomsGroupes.add(groupe, groupe.nom)
    else
      dim typeGg as TypeGroupeT
      typeGg.typeGroupe = groupe.typeGroupe
      dim nouveauNomGroupe as new Collection
      nouveauNomGroupe.add(groupe, groupe.nom)
      typeGg.nomsGroupes = nouveauNomGroupe
      tableauTypesGroupe.add(typeGg, groupe.typeGroupe)
    endif
  next index
  
  ' tri des groupes au sein de chaque type
  for index = 1 to tableauTypesGroupe.count
    tri(tableauTypesGroupe.item(index).nomsGroupes)
  next index

  getEnsemblesParType = tableauTypesGroupe
End Function


'**************************************************************
' Tri la collection dans l'ordre alphabétique du nom du groupe
'**************************************************************
Sub tri(byRef collectionGroupes as new Collection)
  dim nombreEnsembles as integer : nombreEnsembles = collectionGroupes.count
  dim idx, jdx, idxPlusPetit as integer
  dim valPlusPetit, valeurTest, cleSuivante as string
  dim elementPetit as GroupeT
  
  for idx = 1 to nombreEnsembles
    ' recherche de l'élément le plus petit
    idxPlusPetit = idx
    valPlusPetit = lcase(collectionGroupes(idx).nom)
    for jdx = idx + 1 to nombreEnsembles
      valeurTest = lcase(collectionGroupes(jdx).nom)
      if valeurTest < valPlusPetit then
        idxPlusPetit = jdx
        valPlusPetit = valeurTest
      end if
    next jdx
    
    ' déplacement du plus petit élément
    if idxPlusPetit <> idx then
      cleSuivante = collectionGroupes.item(idx).nom ' recherche de la clé suivante pour l'insertion
      elementPetit = collectionGroupes.item(idxPlusPetit)
      collectionGroupes.remove(idxPlusPetit) ' suppression de l'élément trouvé, dont l'indice vient d'augmenter
      collectionGroupes.add(elementPetit, elementPetit.nom, Before:=cleSuivante) ' insertion de l'élément le plus petit
    end if
  next idx
  
end Sub


'******************************************************************
' Affiche les cases à cocher d'un type de groupe
'******************************************************************
Sub afficheToutesCheckBoxDeType(nomGroupeAAfficher as string, byRef idxCheckBox as integer, byRef tableauTypesGroupe as new Collection, absColonne as long)
  dim libelleCase as string, indexGr as integer, groupe as GroupeT
  dim typeGr as TypeGroupeT : typeGr = tableauTypesGroupe(nomGroupeAAfficher)
  for indexGr = 1 to typeGr.nomsGroupes.count
    groupe = typeGr.nomsGroupes.item(indexGr)
    libelleCase = groupe.nom & " (" & groupe.numAdherents.count & " élèves)"
    dlg_Coche(absColonne, 18 + idxCheckBox * 14, 100, 14, "checkBoxEnsemble" & groupe.nom, libelleCase, false)    
    idxCheckBox = idxCheckBox + 1
  next indexGr
End Sub


'************************************************************
' Création d'une boite vide
'   x, y : Position X et Y
'   larg, haut : Largeur et hauteur
'   cTitre : Titre
'************************************************************
Sub dlg_Creation(x as long, y as long, larg as long, haut as long, cTitre as string)
   oDlgModele = createUnoService("com.sun.star.awt.UnoControlDialogModel")
   
   oDlgModele.PositionX = x
   oDlgModele.PositionY = y
   oDlgModele.Width = larg
   oDlgModele.Height = haut
   oDlgModele.Title = cTitre

   oDlgControle = createUnoService("com.sun.star.awt.UnoControlDialog")
   oDlgControle.setModel( oDlgModele )
   
End Sub


'************************************************************
' Création d'une boite à cocher. 
'   x, y : Position X et Y
'   larg, haut : Largeur et hauteur
'   cNom : Nom logique, titre et nom du listener (optionnel)
'************************************************************
Sub dlg_Coche(x as long, y as long, larg as long, haut as long, cNom as string, cLib as string,_
                  Optional bCoche as Boolean                  )
                  ', Optional cNomListener as string 
                  ')
                  
   dim oCocheModele as object
   oCocheModele = oDlgModele.createInstance( "com.sun.star.awt.UnoControlCheckBoxModel" )
   ' Initialize the button model's properties.
   oCocheModele.PositionX = x
   oCocheModele.PositionY = y
   oCocheModele.Width = larg
   oCocheModele.Height = haut
   oCocheModele.Name = cNom
   oCocheModele.Label = cLib
   oCocheModele.State = 0
   If bCoche Then
      oCocheModele.State = 1
   EndIf
   
   oDlgModele.insertByName( cNom, oCocheModele )
   'oCocheControle = oDlgControle.getControl( cNom )

   ' Les boutons doivent avoir une écoute
   ' Création d'une procédure pour recevoir l'événement
   'If IsMissing( cNomListener ) Then
   '   cNomListener = cNom
   'EndIf
   'oActionListener = CreateUnoListener( cNomListener + "_", "com.sun.star.awt.XActionListener" )
   'oCocheControle.addActionListener( oActionListener )

End Sub


'********************************
' Affiche une ligne horizontale
'********************************
Sub afficheLigne(x as long, y as long, larg as long, haut as long, cNom as string)
   dim oLigneModele as object
   oLigneModele = oDlgModele.createInstance( "com.sun.star.awt.UnoControlFixedLineModel" )
   oLigneModele.PositionX = x
   oLigneModele.PositionY = y
   oLigneModele.Width = larg
   oLigneModele.Height = haut
   oDlgModele.insertByName( cNom, oLigneModele )
   
End Sub

'************************************************************
' Création d'un libellé
'   x, y : Position X et Y
'   larg, haut : Largeur et hauteur
'************************************************************
Sub dlg_Libelle(x as long, y as long, larg as long, haut as long, cNom as string, cLib as string )
   
   dim oLibModele as object
   oLibModele = oDlgModele.createInstance( "com.sun.star.awt.UnoControlFixedTextModel" )

   oLibModele.PositionX = x
   oLibModele.PositionY = y
   oLibModele.Width = larg
   oLibModele.Height = haut
   oLibModele.Name = cNom
   oLibModele.Label = cLib
   
   oDlgModele.insertByName( cNom, oLibModele )
   dim oLibControle as object
   oLibControle = oDlgControle.getControl( cNom )

End Sub


'************************************************************
' Création d'un bouton
'   x, y : Position X et Y
'   larg, haut : Largeur et hauteur
'   cTitre : Titre
'   Nom logique
'   nom du listener (optionnel)
'************************************************************
Sub dlg_Bouton(x as long, y as long, larg as long, haut as long, cNom as string, cLib as string,_
                  Optional cNomListener as string )
     
   dim oBoutonModele as object
   oBoutonModele = oDlgModele.createInstance( "com.sun.star.awt.UnoControlButtonModel" )

   oBoutonModele.PositionX = x
   oBoutonModele.PositionY = y
   oBoutonModele.Width = larg
   oBoutonModele.Height = haut
   oBoutonModele.Name = cNom
   oBoutonModele.Label = cLib
   
   oDlgModele.insertByName( cNom, oBoutonModele )
   dim oBoutonControle as object
   oBoutonControle = oDlgControle.getControl( cNom )

   ' Création du listener
   If IsMissing( cNomListener ) Then
      cNomListener = cNom
   EndIf
   dim oActionListener as object
   oActionListener = createUnoListener( cNomListener + "_", "com.sun.star.awt.XActionListener" )
   oBoutonControle.addActionListener( oActionListener )
     
End Sub


'************************************************************
' Affiche la boite
'************************************************************
Sub dlg_Affiche()
   oDlgControle.setVisible( True )
   oDlgControle.execute()
End Sub

'************************************************************
' Ferme la boite
'************************************************************
Sub dlg_Ferme()
   oDlgControle.endExecute()
   oDlgControle.setVisible( False )
End Sub


'************************************************************
' Listener unique pour OK ou Annuler
'************************************************************
Sub ValideChoixEnsembles_actionPerformed(oEve)
  dim checkBoxModel as object
  dim idxColl as integer
  dim groupe as GroupeT
  nbEnsemblesSelectionnes = 0
  for idxColl = 1 to numAdherentsParEnsemble.count
   	groupe = numAdherentsParEnsemble.item(idxColl)
	checkBoxModel = oDlgControle.getControl("checkBoxEnsemble" & groupe.nom).getModel()
	if checkBoxModel.state then
	  ensemblesChoisisTab(nbEnsemblesSelectionnes) = idxColl
	  nbEnsemblesSelectionnes = nbEnsemblesSelectionnes + 1
	end if   	  
  next idxColl
  ' fermeture de la boite de dialogue
  dlg_Ferme()
End Sub


Sub SuiteRecupereEmails_actionPerformed(oEve)
  ' fermeture de la boite de dialogue
   dlg_Ferme()
End Sub

Sub Annule_actionPerformed(oEve)
  annule = true
  ' fermeture de la boite de dialogue
   dlg_Ferme()
End Sub


Sub afficheEnsemblesTrouvesPourDebug(byRef ensemblesTabl() as string, byRef collAdherents() as new Collection, nbAdherents as integer, feuille as object)
  feuille.getCellRangeByName("AT1").string = (nbAdherents & " adhérents")
  dim idxTab as integer
  dim texteAAfficher as string
  dim idxColl as integer
  dim collObtenue as new Collection
  dim valeurNumAdherent as integer
  for idxTab = 0 to ubound(ensemblesTabl)
    texteAAfficher = ensemblesTabl(idxTab) & " : "
    collObtenue = collAdherents(ensemblesTabl(idxTab))
    for idxColl = 1 to collObtenue.count
      valeurNumAdherent = collObtenue.item(idxColl)
      texteAAfficher = texteAAfficher & valeurNumAdherent & " "
    next idxColl
    feuille.getCellRangeByName("AT"&(idxTab+2)).string = texteAAfficher
  next idxTab
  feuille.getCellRangeByName("AT20").string = (collAdherents.count & " ensembles")
End Sub


'************************************************************
'  Récupération des adresses mails des élèves sélectionnés
'  Si un élève est dans plusieurs ensembles, ses adresses ne sont prises qu'une seule fois
'************************************************************
Function recupereEmails(byRef collAdherentsSlectionnes() as new Collection, feuille as object)
  dim indexMax as integer : indexMax = ubound(collAdherentsSlectionnes) ' nombre d'ensembles sélectionnés - 1
  dim numColEmails1 as long : numColEmails1 = feuille.getCellRangeByName("TITRE_ADR_MAIL1").cellAddress.column ' numéro de la colonne 'MAIL1'
  dim numColEmails2 as long : numColEmails2 = numColEmails1 + 1 ' numéro de la colonne 'MAIL2'
  dim indexesCollections(indexMax) as integer ' tableau contenant un numéro d'index pour chaque collection
  dim taillesMaxCollections(indexMax) as integer ' tableau contenant la taille max de chaque collection (donc le nombre d'élèves de cet ensemble)
  dim plusPetiteValeur as integer ' numéro d'adhérent
  dim val as integer
  dim val2 as integer
  dim collTmp as new Collection
  dim idxCTmp as integer
  dim idx0, idx1, idx2, idx3 as integer : idx3 = 0
  dim nbAdherentsSelectionnes as integer : nbAdherentsSelectionnes = 0 
  dim nombreAdherentsEnsembleCourant as integer
  
  ' initialisation des tableaux
  for idx0 = 0 to indexMax
    indexesCollections(idx0) = 1
    nombreAdherentsEnsembleCourant = collAdherentsSlectionnes(idx0).count
    taillesMaxCollections(idx0) = nombreAdherentsEnsembleCourant
    nbAdherentsSelectionnes = nbAdherentsSelectionnes + nombreAdherentsEnsembleCourant
  next idx0

  dim adressesMail(nbAdherentsSelectionnes * 2) as string ' résultat de la fonction : tableau contenant les emails. La taille du tableau est une majoration.
  
  do
    ' parcours des collections, en choisissant le plus petit numéro d'adhérent
    plusPetiteValeur = 9999
    for idx1 = 0 to indexMax
      if indexesCollections(idx1) <= taillesMaxCollections(idx1) then
        collTmp = collAdherentsSlectionnes(idx1)
        idxCTmp = indexesCollections(idx1)
        val = collTmp.item(idxCTmp)
        if val < plusPetiteValeur then
          plusPetiteValeur = val
        endif
      endif
    next idx1
    
    ' lecture des emails de cet élève
    if plusPetiteValeur < 9999 then
      adressesMail(idx3) = rtrim(feuille.getCellByPosition(numColEmails1, plusPetiteValeur).string)
      idx3 = idx3 + 1
      if feuille.getCellByPosition(numColEmails2, plusPetiteValeur).type <> com.sun.star.table.CellContentType.EMPTY then
        adressesMail(idx3) = rtrim(feuille.getCellByPosition(numColEmails2, plusPetiteValeur).string)
        idx3 = idx3 + 1
      endif
    endif
    
    ' pour chaque collection possédant cette valeur, incrément de l'index
    for idx2 = 0 to indexMax
      if indexesCollections(idx2) <= taillesMaxCollections(idx2) then
        if collAdherentsSlectionnes(idx2).item(indexesCollections(idx2)) = plusPetiteValeur then
          indexesCollections(idx2) = indexesCollections(idx2) + 1
        endif
      endif
    next idx2
    
  loop until plusPetiteValeur = 9999

  recupereEmails = adressesMail
end function

'************************************************************
' Restitue les emails en les copiant dans le presse-papier, 
' en plusieurs fois s'il y en a beaucoup
'************************************************************
Sub restitueEmails(byRef emailsSelectionnes() as string)
    dim nbEmailsTotal as integer : nbEmailsTotal = compteElementsTableau(emailsSelectionnes) ' nombre total d'emails correspondant aux ensembles sélectionnés
    dim nbCopiesPressPapierTotal as integer : nbCopiesPressPapierTotal = fix(nbEmailsTotal / NB_EMAILS_MAX) + 1 ' nombre de copies nécessaires dans le presse-papier pour la prise en compte de tous les emails
    dim nbCopiesPressPapierFait as integer : nbCopiesPressPapierFait = 1   
    dim indexEmailsSel as integer
    dim concatenationEmails as string
    dim compteurEmailsPP as integer : compteurEmailsPP = 0 ' nombre d'emails dans le press-papier

    
    ' Parcours des emails pour les copiers dans le presse-papier, tout en limitant leur nombre à NB_EMAILS_MAX
    for indexEmailsSel = 0 to nbEmailsTotal - 1
      concatenationEmails = concatenationEmails & emailsSelectionnes(indexEmailsSel)
      compteurEmailsPP = compteurEmailsPP + 1
      if compteurEmailsPP = NB_EMAILS_MAX and not annule then
        ' copie des adresses dans le presse-papier
        copyToClipboard(concatenationEmails)
        nbCopiesPressPapierFait = nbCopiesPressPapierFait + 1
        if indexEmailsSel + 1 < nbEmailsTotal then
          ' affiche une boite de dialogue indiquant qu'il reste des adresses à récupérer
          afficheBoiteDialogueResteEmails(nbCopiesPressPapierFait, nbCopiesPressPapierTotal, indexEmailsSel, nbEmailsTotal)
          compteurEmailsPP = 0
          concatenationEmails = ""
        endif
      endif
    next indexEmailsSel
    
    if compteurEmailsPP <> NB_EMAILS_MAX and not annule then
      ' copie les dernières adresses dans le presse-papier
      copyToClipboard(concatenationEmails)
    end if 
end sub


'************************************************************
' Compte le nombre d'éléments valorisés d'un tableau.
' Contrainte : les éléments non valorisés sont tous à la fin.
'************************************************************
Function compteElementsTableau(byRef tabEmailsSelectionnes() as string) as integer
  dim tailleTableau as integer : tailleTableau = uBound(tabEmailsSelectionnes)
  dim index as integer
  ' Hypothèse d'optimisation : le tableau contient au moins la moitié de ses cases valorisées
  index = fix(tailleTableau / 2) - 1
  ' Correction si cette hypothèse était fausse :
  if tabEmailsSelectionnes(index) = "" then
    index = -1
  endif
  
  do
    index = index + 1
  loop until index = tailleTableau or tabEmailsSelectionnes(index) = ""

  compteElementsTableau = index
end function

 
'************************************************************
' Copie le texte fourni en paramètre dans le presse-papier 
'************************************************************
Sub copyToClipboard(sText)
  dim oClip, oTR as object
  ' create SystemClipboard instance
  oClip = CreateUnoService( "com.sun.star.datatransfer.clipboard.SystemClipboard")
  oTR = createUnoListener("Tr_", "com.sun.star.datatransfer.XTransferable")
  ' set data
  oClip.setContents( oTR,Null )
  sTxtCString = sText
  ' oClip.flushClipboard() ' does not work
End Sub
 
Function Tr_getTransferData(aFlavor as com.sun.star.datatransfer.DataFlavor)
  If (aFlavor.MimeType = "text/plain;charset=utf-16") Then
    Tr_getTransferData() = sTxtCString
  End If
End Function
 
Function Tr_getTransferDataFlavors()
  Dim aFlavor As new com.sun.star.datatransfer.DataFlavor
  aFlavor.MimeType = "text/plain;charset=utf-16"
  aFlavor.HumanPresentableName = "Unicode-Text"
  Tr_getTransferDataFlavors() = array( aFlavor )
End Function
 
Function Tr_isDataFlavorSupported(aFlavor as com.sun.star.datatransfer.DataFlavor) as Boolean
  If aFlavor.MimeType = "text/plain;charset=utf-16" Then
    Tr_isDataFlavorSupported = true
  Else
    Tr_isDataFlavorSupported = false
  End If
End Function
