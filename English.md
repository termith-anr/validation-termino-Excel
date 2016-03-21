GenereXLSX.pl
=============

Perl script to generate an Excel file for validating new terminological entries in the context of the TermITH project. ([version française](https://github.com/termith-anr/validation-termino-Excel))

### Requirements

- [Perl](https://www.perl.org/) (Version ≥ 5.8.3)

- Perl modules

  - usually part of the Perl distribution:
  
    - Encode
    
    - Getopt::Long
    
    - POSIX
    
  -  to download if necessary from a repository like *[CPAN](http://www.cpan.org/modules/index.html)* :
  
    - [XML::Twig](http://xmltwig.org/xmltwig/)
    
    - [Excel::Writer::XLSX](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/)
    
    - Excel::Writer::XLSX::Utility (included in the "Excel::Writer::XLSX" package)

# Running the programme

## Synopsis

```
GenereXLSX.pl -c candidates -l lexicon -e Excel file
              -u Totem URL ( -a Aroma | -s Smoa )
            [ -r report ] [ -d domain ]
            [ -v 'name of vocabulary' ]
```

## Mandatory options

  - \-c *candidates*     : list of potential terms in TBX format
  
  - \-l *lexicon*        : terminological lexicon of reference in TBX or RDF format
  
  - \-e *Excel file*     : name of the generated Excel file
  
  - \-r *Totem URL*      : URL of the instance of the [Totem server](https://github.com/termith-anr/totem) used to display the occurrences of the potential terms from the studied corpus
  
  - \-a *Aroma*          : file in **ALIGN** format of matchings between potential terms and terms from the reference lexicon obtained with the **AROMA** method
  
  - \-s *Smoa*           : file in **ALIGN** format of matchings between potential terms and terms from the reference lexicon obtained with the **SMOA** method
  
## facultative options

  - \-r *report*        : name of the logfile with extra information on the Excel file generation
  
  - \-d *domain*        : name of the domain to which the terminolical lexicon belongs
  
  - \-v *vocabulary*    : name of the vocabulary, i.e. the terminological lexicon of reference

# Data files

## Potential terms

File in TBX format containing for each potential term an id (*xml:id*), the term in lemmatized form (*\<term\>*), the pilot term (i.e. the most frequent form of the potential term in the corpus) and the list of all the forms found in the corpus (*\<descrip type="formList"\>*). 

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

## Terminological lexicon of reference

File containing the terms of the reference lexicon either in RDF format

```xml
<rdf:Description rdf:about="http://www.termsciences.fr/vocabs/ML/113227">
   <rdf:type rdf:resource="http://www.termsciences.fr/vocabs/ML"/>
   <rdfs:label xml:lang="fr">Anaphore associative</rdfs:label>
   <rdfs:label xml:lang="en">Associative anaphora</rdfs:label>
   <notation xmlns="http://www.w3.org/2004/02/skos/core#">113227</notation>
</rdf:Description>
```

or in TBX format

```xml
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

## Alignments

File in **ALIGN** format of the matchings between potential terms and terms of the terminological lexicon of reference obtained either with the **AROMA** method or the **SMOA** method

Example of an alignment with the AROMA method of the term "**anaphore associative**":

```xml
<map>
   <Cell rdf:about="414607">
      <entity1 rdf:resource='http://www.termsciences.fr/vocabs/CandidatsLinguistique/entry-1092493'/>
      <entity2 rdf:resource='http://www.termsciences.fr/vocabs/ML/113227'/>
      <relation>=</relation>
      <measure rdf:datatype='http://www.w3.org/2001/XMLSchema#float'>0.9277312808321435</measure>
      <alignapilocalns:hasCellStatus xmlns:alignapilocalns="http://www.mondeca.com/system/publishing#">
         http://www.mondeca.com/system/publishing#ToBeReviewed
      </alignapilocalns:hasCellStatus>
   </Cell>
</map>
```


