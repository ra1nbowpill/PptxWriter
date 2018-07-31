# Generation de PowerPoint

## Dépendances

- [pptx](https://pypi.org/project/python-pptx/)
	- Manipulation de fichiers .pptx
- [pd2pptx](https://github.com/robintw/PandasToPowerpoint)
	- Ecrire un dataframe dans un slide powerpoint
	- La version fournie a été modifiée pour les besoin de ce paquet


## Fichier
- util.py : fichier source
- template.pptx : template powerpoint a utiliser
- examples : exemples d'utilisation de la librairie

## Création du template

Le fichier `template.pptx` est un fichier pptx qui ne contient aucune diapositives mais contient des masques de diapositives. Les placeholders utilisés dans les masques pour afficher différents types d'objets powerpoint doivent être inséré dans l'ordre de lecture du masque (si il y a 2 placeholders a droite et à gauche de la diapositive il faut placer le premier placeholder créé à gauche puis le second à droite).
Ce sont ces masques qui seront utilisé dans l'objet PptxWriter du fichier `util.py`.

Pour faire le lien entre les masques et le code il est nécéssaire de définir un dictionnaire python (`layouts`) qui associe a un nom défini par l'utilisateur·ice l'indice du masque correspondant. Les noms des masques peuvent suivre la règle suivante pour en faciliter l'utilisation:
- '2_cha_r' : deux graphiques (chart) horizontaux (row)
- '3_tab_c' : trois tableaux en colonnes
- '3_tab_1_cha_r' : une ligne avec 3 tableau et une seconde avec 1 graphique

## Cas d'utilisation