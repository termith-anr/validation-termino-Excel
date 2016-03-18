GenereXLSX.pl
=============

Script Perl de génération d'un fichier Excel servant à la validation de nouvelles entrées terminologiques dans le cadre du projet TermITH. 

*[En savoir plus sur Termith](http://www.atilf.fr/ressources/termith/)*

### Les prérequis

- [Perl](https://www.perl.org/) (Version ≥ 5.8.3)

- Modules Perl

  - normalement présents dans la distribution Perl :
  
    - Encode
    
    - Getopt::Long
    
    - POSIX
    
  -  à charger depuis le dépôt *[CPAN](http://www.cpan.org/modules/index.html)* :
  
    - [XML::Twig](http://xmltwig.org/xmltwig/)
    
    - [Excel::Writer::XLSX](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/)

# Lancer le programme

## Synopsis

```
GenereXLSX.pl -c candidats -l lexique -e fichier Excel
              -u URL Totem ( -a Aroma | -s Smoa )
            [ -r rapport ] [ -d domaine ]
            [ -v 'nom du vocabulaire' ]
```

## Options obligatoires

  - \-c *candidats*     : liste des candidats termes au format TBX
  
  - \-l *lexique*       : lexique terminologique de référence au format TBX ou RDF
  
  - \-e *fichier excel* : nom du fichier Excel à générer
  
  - \-r *URL Totem*     : URL de l'instance du serveur [Totem](https://github.com/termith-anr/totem) permettant d'afficher les occurrences des candidats termes du corpus étudié
  
  - \-a *Aroma*         : fichier au format **ALIGN** des alignements obtenus entre candidats termes et termes du lexique de référence par la méthode **AROMA**
  
  - \-s *Smoa*          : fichier au format **ALIGN** des alignements obtenus entre candidats termes et termes du lexique de référence par la méthode **SMOA**
  
## Options facultatives

  - \-r *rapport*       : nom du fichier de *log* contenant des informations supplémentaires sur la génération du fichier Excel
  
  - \-d *domaine*       : nom du domaine auquel appartient le lexique terminologique de référence
  
  - \-v *vocabulaire*   : nom du vocabulaire, c'est-à-dire le lexique terminologique de référence

# Fichiers de données

## Candidats termes

Fichier au format TBX contenant notamment pour chaque candidat terme un identifiant (*xml:id*), le terme en forme lemmatisée (*<term>*), le terme pilote (c'est-à-dire la forme la plus fréquente du candidat terme dans le corpus) et la liste de toutes les formes trouvées dans le corpus (*<descrip type="formList">*)

```xml
<termEntry xml:id="TS1.4-entry-1092493">
   <langSet xml:id="langset-1092493" xml:lang="fr">
      <descrip target="langset-4109007" type="termVariant">anaphore associatif de type</descrip>
      <descrip type="nbOccurrences">10</descrip>
      <tig xml:id="term-1092493">
         <term>anaphore associatif</term>
         <termNote type="termPilot">anaphore associative</termNote>
         <termNote type="termType">termEntry</termNote>
         <termNote type="partOfSpeech">noun</termNote>
         <termNote type="termPattern">noun-adjective</termNote>
         <termNote type="termComplexity">multi-word</termNote>
         <descrip type="termSpecificity">4.0876</descrip>
         <descrip type="nbOccurrences">10</descrip>
         <descrip type="relativeFrequency">0.0001</descrip>
         <descrip type="formList">
            [{term="anaphore associative", count=7}, 
            {term="anaphores associatives", count=3}]
         </descrip>
      </tig>
   </langSet>
</termEntry>
```

## Lexique de référence

Fichier contenant les termes du lexique de référence soit au format RDF 

```xml
<rdf:Description rdf:about="http://www.termsciences.fr/vocabs/ML/113227">
   <rdf:type rdf:resource="http://www.termsciences.fr/vocabs/ML"/>
   <rdfs:label xml:lang="fr">Anaphore associative</rdfs:label>
   <rdfs:label xml:lang="en">Associative anaphora</rdfs:label>
   <notation xmlns="http://www.w3.org/2004/02/skos/core#">113227</notation>
</rdf:Description>
```

soit au format TBX

```
<termEntry xmlns="http://www.tbx.org" xml:id="BV.113227">
   <descrip type="originatingDatabaseName">Vocabulaire Linguistique</descrip>
   <admin type="conceptIdentifier">BV.113227</admin>
   <admin type="originatingInstitution">Base FRANCIS</admin>
   <admin type="elementWorkingStatus">consolidatedElement</admin>
   <langSet xml:lang="fr">
      <tig>
         <tei:term>Anaphore associative</tei:term>
         <termNote type="administrativeStatus">preferredTerm</termNote>
         <admin type="termIdentifier">BV.113227.1</admin>
         <admin type="originatingDatabaseName">Vocabulaire Linguistique</admin>
         <admin type="inputDate">2002-01-01</admin>
      </tig>
   </langSet>
   <langSet xml:lang="en">
      <tig>
         <tei:term>Associative anaphora</tei:term>
         <termNote type="administrativeStatus">preferredTerm</termNote>
         <admin type="termIdentifier">BV.113227.2</admin>
         <admin type="originatingDatabaseName">Vocabulaire Linguistique</admin>
         <admin type="inputDate">2002-01-01</admin>
      </tig>
   </langSet>
</termEntry>
```

## Alignement

Fichier au format **ALIGN** des alignements obtenus entre candidats termes et termes du lexique de référence soit par la méthode **AROMA**, soit par la méthode **SMOA**. 

Exemple d'alignement par la méthode AROMA du terme "**anaphore associative**" :

```
<map>
   <Cell rdf:about="414607">
      <entity1 rdf:resource='http://www.termsciences.fr/vocabs/CandidatsLinguistique/entry-1092493'/>
      <entity2 rdf:resource='http://www.termsciences.fr/vocabs/ML/113227'/>
      <relation>=</relation>
      <measure rdf:datatype='http://www.w3.org/2001/XMLSchema#float'>0.9277312808321435</measure>
      <alignapilocalns:hasCellStatus xmlns:alignapilocalns="http://www.mondeca.com/system/publishing#">http://www.mondeca.com/system/publishing#ToBeReviewed</alignapilocalns:hasCellStatus>
   </Cell>
</map>
```


