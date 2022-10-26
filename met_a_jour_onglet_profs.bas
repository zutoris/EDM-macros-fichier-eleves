option explicit ' oblige à la déclaration explicite de toutes les variables du programme

'*************************************
' Déclaration des types personnalisés
'*************************************
type ProfT ' données d'un prof
  nom as string ' nom du prof
  nombreHeures as double ' nombre d'heures du prof (cours individuels uniquement)
end type

'*************************************
' Déclaration des variables globales
'*************************************
private const NUM_COL_NB_HEURES = 1 ' numéro de colonne de l'onglet des profs, dans lequel il faut écrire le nombre d'heures
private document as object ' le fichier Calc
private ongletEleves as object ' onglet des élèves
private ongletProfs as object ' onglet des profs


'************************************************************
'    Programme principal
'************************************************************
Sub Main
  document = ThisComponent ' le fichier Calc
  ongletEleves = document.Sheets(0) ' 1er onglet du tableur
  ongletProfs = document.currentController.activeSheet ' onglet courant : celui des profs
  dim numColProf as long : numColProf = ongletEleves.getCellRangeByName("TITRE_PROF").CellAddress.Column ' numéro de la colonne des professeur
  dim numColDuree as long : numColDuree = ongletEleves.getCellRangeByName("DUREE_COURS").CellAddress.Column ' numéro de la colonne de la durée des cours
  dim profs as new Collection
  dim numAdherent as integer : numAdherent = 0 ' nombre d'adhérents, dont ceux qui sont barrés
  dim nomProf as string
  dim duree as double

  ' parcours de la liste des élèves
  do while ongletEleves.getCellByPosition(0, numAdherent+1).type <> com.sun.star.table.CellContentType.EMPTY
    numAdherent = numAdherent + 1
    if ongletEleves.getCellByPosition(0, numAdherent).charStrikeout <= 0 Then ' exclusion des cellules barrées dans la 1ère colonne

      nomProf = ongletEleves.getCellByPosition(numColProf, numAdherent).string
      if nomProf <> "" then 
        duree = ongletEleves.getCellbyPosition(numColDuree, numAdherent).value
        ajouteDuree(duree, rtrim(nomProf), profs)
      end if

    end if
  loop

  ' écriture du nombre d'heures dans l'onglet des profs
  valoriseColonneNombreHeures(profs)
 
End Sub


'***************************************************************************
' Ajoute la durée du cours au total du prof. Le prof est également ajouté 
' à la collection s'il n'était pas déjà présent.
'***************************************************************************
Sub ajouteDuree(duree as double, nomProfess as string, byRef professeurs as new Collection)
	dim prof as ProfT
	
	if estGroupeDansCollection(professeurs, nomProfess) then
	  ' ce prof est déjà présent, 
	  prof = professeurs(nomProfess)
	  prof.nombreHeures = prof.nombreHeures + duree
	else
	  ' prof non présent, il est ajouté
	  prof.nom = nomProfess
	  prof.nombreHeures = duree
	  professeurs.add(prof, nomProfess)
	endif
end sub



'******************************************************************
' Indique si la collection possède la clé fournie. Retourne 'true' si oui, 'false' sinon.
'******************************************************************
Function estGroupeDansCollection(byRef pColl as object, pCle as string) as boolean
	dim element as variant, existe as boolean
	on local error goto ErrHandler ' gérer l'erreur
		existe = false
		element = pColl(pCle) ' si pCle n'existe pas -> erreur
		existe = true
	ErrHandler:
		'ne rien faire
	estGroupeDansCollection = existe
End Function

'***************************************************************************
' écriture du nombre d'heures dans l'onglet des profs
'***************************************************************************
Sub valoriseColonneNombreHeures(byRef professeurs as new Collection)
																																																								  
  dim numLigneProf as integer ' numéro de ligne du prof
  dim profConcerne as ProfT
  dim trouve as boolean ' indique si la ligne du prof a été trouvée
  
  for each profConcerne in professeurs
    ' recherche la ligne du prof
    numLigneProf = 1
    trouve = false
    do while ongletProfs.getCellByPosition(0, numLigneProf).type <> com.sun.star.table.CellContentType.EMPTY
      if profConcerne.nom = ongletProfs.getCellByPosition(0, numLigneProf).string then
        ongletProfs.getCellByPosition(NUM_COL_NB_HEURES, numLigneProf).value = profConcerne.nombreHeures
        trouve = true
      end if
      numLigneProf = numLigneProf + 1
    loop
    
    ' si aucune ligne ne correspond au prof, une nouvelle ligne est ajoutée
    if not trouve then
      ongletProfs.getCellByPosition(0, numLigneProf).string = profConcerne.nom
																				   
      ongletProfs.getCellByPosition(NUM_COL_NB_HEURES, numLigneProf).value = profConcerne.nombreHeures
    end if
  next profConcerne
end sub
