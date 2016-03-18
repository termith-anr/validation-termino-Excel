#!/usr/bin/perl


# Déclaration des pragmas
use strict;
# use warnings;
use utf8;
use open qw/:std :utf8/;

# Appel des modules externes
use POSIX;
use Encode;
use Getopt::Long;
use XML::Twig;

# Recherche du nom du programme
my ($programme) = $0 =~ m|^(?:.*/)?(.+)|;

# Initialisation des variables globales
# nécessaires à la lecture des options
my $aroma      = "";
my $smoa       = "";
my $candidats  = "";
my $domaine    = "";
my $lexique    = "";
my $log        = "";
my $methode    = "";
my $sortie     = "";
my $url_totem  = "";
my $vocab      = "";

eval	{
	$SIG{__WARN__} = sub {usage(1);};
	GetOptions(
		"aroma=s"        => \$aroma,
		"candidats=s"    => \$candidats,
		"excel=s"        => \$sortie,
		"lexique=s"      => \$lexique,
		"domaine=s"      => \$domaine,
		"rapport=s"      => \$log,
		"smoa=s"         => \$smoa,
		"url_totem=s"    => \$url_totem,
		"vocabulaire=s"  => \$vocab,
		);
	};
$SIG{__WARN__} = sub {warn $_[0];};

# Vérification de la présence des options obligatoires
if ( not $candidats or not ( $aroma or $smoa ) or not $lexique or not $url_totem ) {
	usage(2);
	}

# Quelques vérifications sur le fichier des termes candidats (en TBX)
if ( not -f $candidats ) {
	print STDERR "$programme : fichier TBX \"$candidats\" absent\n";
	exit 3;
	}
elsif ( not -r $candidats ) {
	print STDERR "$programme : lecture du fichier TBX \"$candidats\" impossible\n";
	exit 4;
	}

# Ouverture du fichier log ou redirection vers "/dev/null"
if ( $log ) {
	open(LOG, ">:utf8", $log) or die "$!,";
	}
else	{
	open(LOG, ">:utf8", "/dev/null") or die "$!,";
	}

# Initialisation des varaibles globales nécessaires 
# au traitement du fichier TBX
my $lemme  = "";
my $pilote = "";
my $pos    = "";
my @messages = ();
my %id = ();
my %candidats = ();
my %synonymes = ();

# Initialisation du parseur et ...
my $twig = XML::Twig->new( 
			   twig_roots => {
				'termEntry' => 1,
				},
			   twig_handlers => {
				'termEntry'                     => \&termEntry,
				'term'                          => sub {$lemme = $_->text;},
				'termNote[@type="termPilot"]'   => sub {$pilote = $_->text;},
				'termNote[@type="termPattern"]' => sub {$pos = $_->text;},
				'descrip[@type="formList"]'     => \&descrip,
				},
			  );

# ... parsage du fichier
$twig->parsefile("$candidats");
$twig->purge;

# Initialisation des variables globales nécessaires 
# au traitement du fichier d'identifiants du lexique
my $termino  = "Terminologie";
my %synonLex = ();

# Initialisation des varaibles globales nécessaires 
# au traitement du lexique (fichier TBX ou RDF)
my $langue     = "";
my $note       = "";
my $termeLex   = "";
my @associes   = ();
my @english    = ();
my @generiques = ();
my @synonymes  = ();
my @termes     = ();
# my %id         = ();	# Liste des identifiants d'un terme pour trouver les doublons
my %note       = ();	# Note d'application
my %associe    = ();	# Terme(s) associé(s)
my %english    = ();	# Traduction(s) anglaise(s)
my %generique  = ();	# Terme(s) générique(s)
my %specifique = ();	# Terme(s) spécifique(s)
my %synonyme   = ();	# Terme(s) synonyme(s)
my %termeLex   = ();	# Terme préférentiel pour un numéro (identifiant)

my @noms = ();
if ( $lexique =~ /^(\w+(\W\w+)*)=(.+)/ ) {
	push(@noms, $1);
	$lexique = $3;
	}
if ( $lexique =~ /\.tbx\z/ ) {
	traiteTBX($lexique);
	}
elsif ( $lexique =~ /\.rdf\z/ ) {
	traiteRDF($lexique);
	}
else	{
	die "Mauvais format de fichier pour le lexique \"$lexique\"\n";
	}

if ( @noms ) {
	my $nom = shift(@noms);
	if ( length($nom) < 32 ) {
		$termino = $nom;
		foreach my $nom (@noms) {
			if ( length($termino) + length($nom) > 31 ) {
				last;
				}
			else	{
				$termino .= "+$nom";
				}
			}
		}
	}

# Initialisation des varaibles globales nécessaires 
# au traitement du fichier d'alignement 
my $ent1 = "";
my $ent2 = "";
my $score = 0.0;
my $relation = "";

# D'abord le fichier "Aroma"
my %aroma = ();
my $align = \%aroma;
traite($aroma) if $aroma;

# Puis le fichier "Smoa"
my %smoa = ();
$align = \%smoa;
traite($smoa) if $smoa;

# Calculs
my %nb    = ();
my %align = ();
my %score = ();
my %label = ();
my %lemme = ();
my %pos   = ();
foreach my $entity (sort {lc($a) cmp lc($b)} keys %aroma) {
	my $num = $entity;
	if ( not $candidats{$num} ) {
		if ( $candidats{"TS2.0-$num"} ) {
			$num = "TS2.0-$num";
			}
		elsif ( $candidats{"TS1.4-$num"} ) {
			$num = "TS1.4-$num";
			}
		else	{
			next;
			}
		}
	$lemme = $candidats{$num}{'lemme'};
	$pilote = $candidats{$num}{'pilote'};
	$pos = $candidats{$num}{'pos'};
	my @synonymes = sort {$candidats{$num}{'forme'}{$b} <=> $candidats{$num}{'forme'}{$a} or
	                      length($a) <=> length($b) or $a cmp $b} keys %{$candidats{$num}{'forme'}};
	$label{$num} = $pilote;
	$score{$num} = 0.0;
	$lemme{$num} = $lemme;
	$pos{$num}   = $pos;
	my @matchs = keys %{$aroma{$entity}};
	foreach my $match (@matchs) {
		my ($relation, $score) = split(/:/, $aroma{$entity}{$match});
		$score{$num} = $score if $score > $score{$num};
		my $exact = compare($match, $lemme, $pilote, @synonymes);
		push(@{$align{$num}}, "$score\tAroma\t$pilote\t$match\t$relation\t$exact");
		}
	if ( $smoa{$entity} ) {
		@matchs = keys %{$smoa{$entity}};
		foreach my $match (@matchs) {
			my ($relation, $score) = split(/:/, $smoa{$entity}{$match});
			$score{$num} = $score if $score > $score{$num};
			my $exact = compare($match, $lemme, $pilote, @synonymes);
			push(@{$align{$num}}, "$score\tSmoa\t$pilote\t$match\t$relation\t$exact");
			}
		$nb{'Smoa'} ++;
		}
	delete $candidats{$num};
	$nb{'Aroma'} ++;
	}

foreach my $entity (sort {lc($a) cmp lc($b)} keys %smoa) {
	my $num = $entity;
	if ( not $candidats{$num} ) {
		if ( $candidats{"TS2.0-$num"} ) {
			$num = "TS2.0-$num";
			}
		elsif ( $candidats{"TS1.4-$num"} ) {
			$num = "TS1.4-$num";
			}
		else	{
			next;
			}
		}
	$lemme = $candidats{$num}{'lemme'};
	$pilote = $candidats{$num}{'pilote'};
	$pos = $candidats{$num}{'pos'};
	my @synonymes = sort {$candidats{$num}{'forme'}{$b} <=> $candidats{$num}{'forme'}{$a} or
	                      length($a) <=> length($b) or $a cmp $b} keys %{$candidats{$num}{'forme'}};
	if ( $lemme ne $pilote ) {
		foreach my $synonyme (@synonymes) {
			if ( $synonyme eq $lemme ) {
				$pilote = $lemme;
				last;
				}
			}
		}
	$label{$num} = $pilote;
	$score{$num} = 0.0;
	$lemme{$num} = $lemme;
	$pos{$num}   = $pos;
	my @matchs = keys %{$smoa{$entity}};
	foreach my $match (@matchs) {
		my ($relation, $score) = split(/:/, $smoa{$entity}{$match});
		$score{$num} = $score if $score > $score{$num};
		my $exact = compare($match, $lemme, $pilote, @synonymes);
		push(@{$align{$num}}, "$score\tSmoa\t$pilote\t$match\t$relation\t$exact");
		}
	delete $candidats{$num};
	$nb{'Smoa'} ++;
	}

foreach my $entity (sort keys %candidats) {
	$lemme = $candidats{$entity}{'lemme'};
	$pilote = $candidats{$entity}{'pilote'};
	$pos = $candidats{$entity}{'pos'};
	my @synonymes = sort {$candidats{$entity}{'forme'}{$b} <=> $candidats{$entity}{'forme'}{$a} or
	                      length($a) <=> length($b) or $a cmp $b} keys %{$candidats{$entity}{'forme'}};
	if ( $lemme ne $pilote ) {
		foreach my $synonyme (@synonymes) {
			if ( $synonyme eq $lemme ) {
				$pilote = $lemme;
				last;
				}
			}
		}
	$label{$entity} = $pilote;
	$score{$entity} = 0.0;
	$lemme{$entity} = $lemme;
	$pos{$entity}   = $pos;
	@{$align{$entity}} = ("0.0\t\t$pilote\t\t\t0");
	delete $candidats{$entity};
	}

# Vérification du nom du fichier de sortie
if ( $sortie !~ /\.xlsx\z/ ) {
	if ( $sortie =~ /\.\w+\z/ ) {
		$sortie =~ s/\.\w+\z/.xslx/;
		}
	else	{
		$sortie .= ".xlsx";
		}
	}

# Initialisation des variables globales nécéssaires à 
# la génération du fichier Excel
my $cell  = "";		# Notation 'A1' pour une cellule
my $debut = "";		# Notation 'A1' pour le début d'un groupe de cellules
my $fin   = "";		# Notation 'A1' pour la fin d'un groupe de cellules
my %couleur = (
	'BrunClair'     => [196, 189, 151],
	'BrunPale'      => [221, 217, 196],
	'JaunePale'     => [255, 255, 204],
	'SaumonPale'    => [253, 233, 217],
	'VertClair'     => [196, 215, 155],
	'VertPale'      => [196, 238, 196],
	);

# Génération d'un fichier Excel
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;

# Création d'un nouveau fichier Excel 2010+
my $excel = Excel::Writer::XLSX->new($sortie);
if ( not defined $excel ) {
	print STDERR "$programme : Impossible de créer le fichier Excel \"$sortie\"\n";
	print STDERR "(il est possible que ce fichier existe déjà et soit ouvert sous Excel)\n";
	exit 5;
	}

# Création de quelques formats
my %format = ();
my $format = $excel->add_format(
			bold        => 1,
			top         => 1,
			top_color   => '#808080',
			valign      => 'top',
			);
$format{'BoldTop1'} = $format;	# format0

$format = $excel->add_format(
			bold        => 1,
			);
$format{'Bold'} = $format;	# format0

$format = $excel->add_format(
			bold        => 1,
			font_script => 1,
			);
$format{'BoldSup'} = $format;	# format0

$format = $excel->add_format(
			font_script => 1,
			);
$format{'Sup'} = $format;	# format0

$format = $excel->add_format(
			bold        => 1,
			font_script => 2,
			);
$format{'BoldSub'} = $format;	# format0

$format = $excel->add_format(
			font_script => 2,
			);
$format{'Sub'} = $format;	# format0

$format = $excel->add_format(
			top         => 1,
			valign      => 'top',
			top_color   => '#808080',
			);
$format{'Top1'} = $format;	# format0

$format = $excel->add_format(
			text_wrap   => 1,
			top         => 1,
			top_color   => '#808080',
			valign      => 'top',
			);
$format{'WrapTop1'} = $format;	# format0

$format = $excel->add_format(
			bold        => 1,
			size        => 12,
			);
$format{'Bold12'} = $format;	# format0

$format = $excel->add_format(
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			);
$format{'BoldCenter'} = $format;

$format = $excel->add_format(
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			top         => 1,
			top_color   => '#808080',
			);
$format{'BoldCenterTop1'} = $format;

$format = $excel->add_format(
			color       => '#9C0006',
			bg_color    => '#FFC7CE',
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			);
$format{'BoldCenterRed'} = $format;

$format = $excel->add_format(
			color       => '#9C0006',
			bg_color    => '#FFC7CE',
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			top         => 1,
			top_color   => '#808080',
			);
$format{'BoldCenterRedTop1'} = $format;

$format = $excel->add_format(
			color       => '#9C6500',
			bg_color    => '#C4EEC4',
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			);
$format{'BoldCenterGreen'} = $format;

$format = $excel->add_format(
			color       => '#9C6500',
			bg_color    => '#C4EEC4',
			align       => 'center',
			valign      => 'vcenter',
			bold        => 1,
			top         => 1,
			top_color   => '#808080',
			);
$format{'BoldCenterGreenTop1'} = $format;

$format = $excel->add_format(
			text_wrap   => 1,
			align       => 'center',
			valign      => 'vcenter',
			size        => 14,
			bold        => 1,
			);
$format->set_bg_color(sprintf("#%02X%02X%02X", @{$couleur{'BrunPale'}}));
$format{'BoldWrapCenterBrunPale14'} = $format;

$format = $excel->add_format(
			text_wrap   => 1,
			align       => 'center',
			valign      => 'vcenter',
			size        => 14,
			bold        => 1,
			);
$format->set_bg_color(sprintf("#%02X%02X%02X", @{$couleur{'BrunClair'}}));
$format{'BoldWrapCenterBrunClair14'} = $format;	# format1b

$format = $excel->add_format(
			num_format  => '0.000',
			align       => 'center',
			);
$format{'CenterNum3d'} = $format;

$format = $excel->add_format(
			num_format  => '0.00',
			align       => 'center',
			);
$format{'CenterNum2d'} = $format;

$format = $excel->add_format(
			num_format  => '0.00',
			align       => 'center',
			top         => 1,
			top_color   => '#808080',
			);
$format{'CenterNum2dTop1'} = $format;

$format = $excel->add_format(
			num_format  => '0.0 %',
			);
$format{'Percent1d'} = $format;

$format = $excel->add_format(
			align       => 'center',
			);
$format{'Center'} = $format;

$format = $excel->add_format(
			align       => 'center',
			top         => 1,
			top_color   => '#808080',
			);
$format{'CenterTop1'} = $format;

$format = $excel->add_format(
			align       => 'center',
			italic      => 1,
			);
$format{'CenterItalic'} = $format;

$format = $excel->add_format(
			valign      => 'vcenter',
			);
$format{'VCenter'} = $format;

$format = $excel->add_format(
			valign      => 'vcenter',
			top         => 1,
			top_color   => '#808080',
			);
$format{'VCenterTop1'} = $format;

$format = $excel->add_format(
			align       => 'left',
			);
$format{'Left'} = $format;

$format = $excel->add_format(
			align       => 'left',
			italic      => 1,
			);
$format{'LeftItalic'} = $format;

foreach my $couleur (keys %couleur) {
	my $valeur = sprintf("#%02X%02X%02X", @{$couleur{$couleur}});
	$format = $excel->add_format(
			bg_color    => $valeur,
			pattern     => 1,
			num_format  => '0.00',
			align       => 'center',
			);
	$format{"CenterNum2d$couleur"} = $format;
	$format = $excel->add_format(
			bg_color    => $valeur,
			pattern     => 1,
			num_format  => '0.00',
			align       => 'center',
			top         => 1,
			top_color   => '#808080',
			);
	$format{"CenterNum2d${couleur}Top1"} = $format;
	}

# Cas du format d'un hyperlien
$format = $excel->add_format( 
			color       => 'blue', 
			underline   => 1,
			);
$format{'UrlNormal'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			underline   => 1,
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlNormalTop1'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			underline   => 1,
			valign      => 'vcenter',
			);
$format{'UrlNormalVCenter'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			underline   => 1,
			valign      => 'vcenter',
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlNormalVCenterTop1'} = $format;

$format = $excel->add_format( 
			bg_color    => '#90EE90',
			color       => 'blue', 
			underline   => 1,
			);
$format{'UrlFondVert'} = $format;

$format = $excel->add_format( 
			bg_color    => '#90EE90',
			color       => 'blue', 
			underline   => 1,
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlFondVertTop1'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			);
$format{'UrlInterneBleu'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			font_script => 1,
			);
$format{'BleuSup'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			font_script => 2,
			);
$format{'BleuSub'} = $format;

$format = $excel->add_format( 
			color       => 'blue', 
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlInterneBleuTop1'} = $format;

$format = $excel->add_format( 
			color       => 'green', 
			);
$format{'UrlInterneVert'} = $format;

$format = $excel->add_format( 
			color       => 'green', 
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlInterneVertTop1'} = $format;

$format = $excel->add_format( 
			color       => 'green', 
			font_script => 1,
			);
$format{'VertSup'} = $format;

$format = $excel->add_format( 
			color       => 'green', 
			font_script => 2,
			);
$format{'VertSub'} = $format;

$format = $excel->add_format( 
			color       => 'brown', 
			);
$format{'UrlInterneBrun'} = $format;

$format = $excel->add_format( 
			color       => 'brown', 
			top         => 1,
			top_color   => '#808080',
			);
$format{'UrlInterneBrunTop1'} = $format;

$format = $excel->add_format( 
			color       => 'brown', 
			font_script => 1,
			);
$format{'BrunSup'} = $format;

$format = $excel->add_format( 
			color       => 'brown', 
			font_script => 2,
			);
$format{'BrunSub'} = $format;

# Ajout d'une feuille de calcul : les paramètres
my $param = $excel->add_worksheet('Paramètres');
my $feuille = $param;

# 
$feuille->set_column(2, 2, 28);
$feuille->set_column(5, 5, 5);
$feuille->set_column(6, 6, 14);
$feuille->set_column(9, 9, 5);
$feuille->set_column(10, 10, 12);
$feuille->set_column(13, 13, 5);
$feuille->set_column(14, 14, 32);

$feuille->write(1, 1, "Domaine traité :", $format{'Bold12'});
if ( $domaine ) {
	$feuille->write(2, 2, decode_utf8($domaine));
	}
else	{
	$feuille->write(2, 2, "<Préciser le nom du domaine>");
	}

# my $np = keys %candidats;
$feuille->write(4, 1, "Nombre de termes candidats :", $format{'Bold12'});
$feuille->write(5, 2, scalar keys %align, $format{'Center'});

my $np = 6;	# Numéro de ligne courante sur la feuille "paramètres"

$feuille->write(++ $np, 1, "Ressources terminologique :", $format{'Bold12'});
if ( $vocab ) {
	$feuille->write(++ $np, 2, decode_utf8($vocab));
	}
else	{
	$feuille->write(++ $np, 2, "<Préciser le nom de la ressource>", $format{'Center'});
	}
$np ++;
$feuille->write(++ $np, 1, "Nombre de termes contrôlés :", $format{'Bold12'});
$feuille->write(++ $np, 2, scalar keys %termeLex, $format{'Center'});

$np ++;
$feuille->write(++ $np, 1, "Méthode d'alignement :", $format{'Bold12'});
$feuille->write(++ $np, 2, "Aroma : ", $format{'Center'});
my $texte = $nb{'Aroma'} > 1 ? sprintf("%d termes alignés", $nb{'Aroma'}) : 
                               sprintf("%d terme aligné", $nb{'Aroma'});
$feuille->write($np, 3, $texte);
if ( $smoa ) {
	$feuille->write(++ $np, 2, "Smoa : ", $format{'Center'});
	$texte = $nb{'Smoa'} > 1 ? sprintf("%d termes alignés", $nb{'Smoa'}) : 
                                    sprintf("%d terme aligné", $nb{'Smoa'});
	$feuille->write($np, 3, $texte);
	}

$np ++;

# Avant de créer la page de résultat, il faut calculer  
# la position des différents termes dans la page dédié 
# au lexique pour pouvoir connaitre la destination des 
# liens interpages
my $base     = 0;
my $nbCells  = 0;
my $nbURLs   = 0;
my $nbTermes = 0;
my %nbCells  = ();
my %nbURLs   = ();
my %position = ();

# Listage des doublons dans le fichier log
my %num = ();
foreach my $terme (keys %termeLex) {
	$num{$terme}{$termeLex{$terme}} ++;
	}
foreach my $terme (sort {lc($a) cmp lc($b)} keys %num) {
	my @tmp = sort {$a <=> $b} keys %{$num{$terme}};
	if ( $#tmp > 0 ) {
		print LOG "Doublon :\t$terme\t";
		print LOG join(", ", sort {$a <=> $b} @tmp), "\n";
		}
	}

# Calcul du nombre de lignes (cellules) et d'URLs par terme
foreach my $terme (@termes) {
	$nbCells = 1;
	$nbURLs  = 0;
	my $tmp = dedoublonne(@{$english{$terme}});
	if ( $nbCells < $tmp ) {
		$nbCells = $tmp;
		}
	$tmp = dedoublonne(@{$synonyme{$terme}});
	if ( $nbCells < $tmp ) {
		$nbCells = $tmp;
		}
	$tmp = dedoublonne(@{$specifique{$terme}});
	if ( $nbCells < $tmp ) {
		$nbCells = $tmp;
		}
	$nbURLs += $tmp;
	$tmp = dedoublonne(@{$generique{$terme}});
	if ( $nbCells < $tmp ) {
		$nbCells = $tmp;
		}
	$nbURLs += $tmp;
	$tmp = dedoublonne(@{$associe{$terme}});
	if ( $nbCells < $tmp ) {
		$nbCells = $tmp;
		}
	$nbURLs += $tmp;
	if ( defined $nbCells{$terme} ) {
		$nbCells{$terme} = $nbCells if $nbCells{$terme} < $nbCells;
		}
	else	{
		$nbCells{$terme} = $nbCells;
		}
	if ( defined $nbURLs{$terme} ) {
		$nbURLs{$terme} = $nbURLs if $nbURLs{$terme} < $nbURLs;
		}
	else	{
		$nbURLs{$terme} = $nbURLs;
		}
	}

$nbCells = $nbURLs = 0;
# Calcul du nombre réel de termes et d'URLs
foreach my $terme (keys %nbCells) {
	$nbTermes ++;
	$nbCells += $nbCells{$terme};
	if ( defined $nbURLs{$terme} ) {
		$nbURLs += $nbURLs{$terme};
		}
	}

# Test du nombre d'URLs 
my $nbPages  = 1;
if ( $nbURLs ) {
	$nbPages  = floor((($nbURLs - 1)/65530) + 1);
	}
my $nbLignes = floor(($nbCells/$nbPages) + 1);
print STDERR "Nombre d’URLs : $nbURLs\n";
# die "Nombre d’URLs trop important\n" if $nbURLs > 65530;

my $page = $termino;
if ( $nbPages > 1 ) {
	$page = substr($termino, 0, 31 - length($nbPages)) . "1";
	}

$nbCells = $nbURLs  = 0;
foreach my $terme (@termes) {
	next if $position{$terme};
	$base ++;
	$position{$terme} = "$page!" . xl_rowcol_to_cell($base, 1);
	$nbCells += $nbCells{$terme};
	if ( defined $nbURLs{$terme} ) {
		$nbURLs += $nbURLs{$terme};
		}
	if ( $nbURLs > 65530 ) {
		$page ++;
		$base = 1;
		$position{$terme} = "$page!" . xl_rowcol_to_cell(1, 1);
		$nbCells = $nbURLs  = 0;
		}
	elsif ( $nbCells >= $nbLignes ) {
		$page ++;
		$base = 0;
		$nbCells = $nbURLs  = 0;
		}
	else	{
		$base += $nbCells{$terme} - 1;
		}
	}

# Ajout d'une feuille de calcul : les résultats
my $result = $excel->add_worksheet('Résultat');
$feuille = $result;
$feuille->activate();

# Réglage de la taille des colonnes et de la première ligne
$feuille->set_column(0, 0, 32);
$feuille->set_column(1, 1, 32, undef, 1, 1, 1);
$feuille->set_column(2, 2, 22);
$feuille->set_column(3, 4, 40);
$feuille->set_column(5, 5, 3, $format{'BoldCenter'});
$feuille->set_column(6, 6, 32, undef, 1, 1);
$feuille->set_column(7, 7, 12, undef, 1, 1);
$feuille->set_column(8, 8, 15, $format{'Center'}, 1, 1);
$feuille->set_column(9, 9, 3, $format{'BoldCenter'}, 1, 1);
$feuille->set_column(10, 10, 30, undef, 1, 1);
$feuille->set_column(11, 11, 40, undef, 1, 1, 1);
$feuille->set_column(13, 13, 30, undef, 1);
$feuille->set_row(0, 40);
# $feuille->set_row(1, 40);

my $nc = 0;	# Numéro de colonne courante
my $nr = 0;	# Numéro de ligne courante sur la feuille "résultats"

my @intitules = (
		'Terme candidat TermITH',				# col.  0	col.  0
#		'Forme lemmatisée',					# col.  1	supprimée
		'Étiquette PoS',					# col.  2	col.  1
		'Validation',						# col.  3	col.  2
		'Commentaire',						# col. 13	col.  3
		'Contexte',						# nouvelle col.	col.  4
		' ',							# col.  4	col.  5
		'Terme contrôlé de la ressource terminologique',	# col.  5	col.  6
		'Score',						# col.  6	col.  7
#		'Méthode',						# col.  7	# supprimée
#		'Relation',						# col.  8	suprimée
		'Alignement validé',					# col.  9	col.  8
		' ',							# col. 10	col.  9
		'Proposition d’enrichissement',				# col. 11	col. 10
		'Terme à enrichir',					# col. 12	col. 11
#		' ',							# nouvelle col.	col. 13
		);
$format = $format{'BoldWrapCenterBrunPale14'};
foreach my $intitule (@intitules) {
	$feuille->write_string($nr, $nc ++, $intitule, $format);
	}

# Maintien de l'entête (première rangée) 
$feuille->freeze_panes(1, 0);

# Est ce que l'URL de Totem se termine par "search" ?
if ( $url_totem =~ m|((https?://\w+(\.\w+)*(:\d+)?)/search)/?\z|o ) {
	$url_totem = $1;
	}
elsif ( $url_totem =~ m|(https?://\w+(\.\w+)*(:\d+)?)/?\z|o ) {
	$url_totem = $1 . "/search";
	}
my $url_lex = "internal:";

# Écriture des résultats
# foreach my $entity (sort {$score{$b} <=> $score{$a} or $label{$a} cmp $label{$b}} keys %align) {
foreach my $entity (sort {lc($label{$a}) cmp lc($label{$b})} keys %align) {
	my ($num) = $entity =~ /entry-(\d+)\z/o;
	if ( $entity =~ /^TS(\d\.\d)-entry-\d+\z/o ) {
		$num = $1 . "-$num";
		}

	my @alignements = tri(@{$align{$entity}});
	my $nb = $#alignements + 1;
	my $nf = $nr + 2;		# rangée de référence au cas où on a plusieurs alignements
	foreach my $alignement ( @alignements ) {
		$nr ++;
		$format = "";
		my ($score, $methode, $pilote, $match, $relation, $exact) = split(/\t/, $alignement);
		my $url_pilote = encode_utf8("\u$pilote");
		$url_pilote =~ s/([^A-Za-z0-9_\n\-+])/sprintf("%%%02X",ord($1))/eg;
		if ( $nb == 1 ) {
			$feuille->write_url($nr, 0, "$url_totem/$num&$url_pilote", $format{'UrlNormalTop1'}, $pilote);
			if ( $pos{$entity} ) {
				$feuille->write_string($nr, 1, $pos{$entity}, $format{'Top1'});
				}
			# Pour obtenir la bordure supérieure
			$feuille->write($nr, 2, undef, $format{'Top1'});
			$feuille->write($nr, 3, undef, $format{'Top1'});
			$feuille->write_string($nr, 4, "", $format{'Top1'});
			$feuille->write_string($nr, 8, "", $format{'CenterTop1'});
			$feuille->write_string($nr, 10, "", $format{'Top1'});
			$feuille->write_string($nr, 11, "", $format{'Top1'});

			$cell = xl_rowcol_to_cell($nr, 5);
			$debut = xl_rowcol_to_cell($nr, 8);
			$feuille->write_formula($nr, 9, "=IF($cell=\"☓\", \"✕\", IF($debut=\"oui\", \"☓\", IF($debut=\"non\", \"⇒\", \"\")))", $format{'Top1'});
			$cell = xl_rowcol_to_cell($nr, 9);
			my $next = xl_rowcol_to_cell($nr, 10);
			$feuille->write_formula($nr, 13, "=IF($cell=\"⇒\", $next, \"\")", $format{'Top1'});
			}
		elsif ( $nb > 1 ) {
			my $dr = $nr + $nb - 1;		# Dernière Rangée
			$feuille->merge_range_type('url', $nr, 0, $dr, 0, "$url_totem/$num&$url_pilote", $format{'UrlNormalVCenterTop1'}, $pilote);
			$feuille->merge_range($nr, 2, $dr, 2, undef, $format{'VCenterTop1'});
			if ( $pos{$entity} ) {
				$feuille->merge_range_type('string', $nr, 1, $dr, 1, $pos{$entity}, $format{'VCenterTop1'});
				}
			# Pour obtenir la bordure supérieure
			$feuille->write($nr, 3, undef, $format{'Top1'});
			$feuille->write_string($nr, 4, "", $format{'Top1'});
			$feuille->write_string($nr, 8, "", $format{'CenterTop1'});
			$feuille->write_string($nr, 11, "", $format{'Top1'});

			$cell = xl_rowcol_to_cell($nr, 5);
			$debut = xl_rowcol_to_cell($nr, 8);
			$fin = xl_rowcol_to_cell($dr, 8);
			$feuille->merge_range($nr, 9, $dr, 9, "=IF($cell=\"☓\", \"✕\", IF(COUNTIF($debut:$fin, \"oui\"), \"☓\", IF(COUNTIF($debut:$fin, \"non\")=$nb, \"⇒\", \"\")))", $format{'BoldCenterTop1'});
			$cell = xl_rowcol_to_cell($nr, 9);
			$feuille->merge_range($nr, 10, $dr, 10, undef, $format{'VCenterTop1'});
			$cell = xl_rowcol_to_cell($nr, 9);
			my $next = xl_rowcol_to_cell($nr, 10);
			$feuille->merge_range($nr, 13, $dr, 13, "=IF($cell=\"⇒\", $next, \"\")", $format{'VCenterTop1'});
			$nb = 0;
			}
		my $ns = $nr + 1;
		if ( $nb == 0 ) {
			if ( $nr < $nf ) {	# Première occurrence
				$feuille->write_formula($nr, 5, "=IF(C$nf=\"accepté (1 terme)\", \"⇒\", IF(C$nf=\"accepté (2 termes)\", \"⇒\", IF(ISTEXT(C$nf)=TRUE, \"☓\", \"\")))", $format{'BoldCenterTop1'});
				}
			else	{
				$feuille->write_formula($nr, 5, "=IF(C$nf=\"accepté (1 terme)\", \"⇒\", IF(C$nf=\"accepté (2 termes)\", \"⇒\", IF(ISTEXT(C$nf)=TRUE, \"☓\", \"\")))");
				}
			}
		else	{
			$feuille->write_formula($nr, 5, "=IF(C$ns=\"accepté (1 terme)\", \"⇒\", IF(C$ns=\"accepté (2 termes)\", \"⇒\", IF(ISTEXT(C$ns)=TRUE, \"☓\", \"\")))", $format{'BoldCenterTop1'});
			}
		if ( $score > 0.90 ) {
			if ( $nr < $nf ) {	# Première occurrence
				$format = $format{'CenterNum2dVertPaleTop1'};
				}
			else	{
				$format = $format{'CenterNum2dVertPale'};
				}
			}
		elsif ( $score > 0.80 ) {
			if ( $nr < $nf ) {	# Première occurrence
				$format = $format{'CenterNum2dJaunePaleTop1'};
				}
			else	{
				$format = $format{'CenterNum2dJaunePale'};
				}
			}
		elsif ( $score > 0.70 ) {
			if ( $nr < $nf ) {	# Première occurrence
				$format = $format{'CenterNum2dSaumonPaleTop1'};
				}
			else	{
				$format = $format{'CenterNum2dSaumonPale'};
				}
			}
		else	{
			if ( $nr < $nf ) {	# Première occurrence
				$format = $format{'CenterNum2dTop1'};
				}
			else	{
				$format = $format{'CenterNum2d'};
				}
			}
		$feuille->write_number($nr, 7, $score, $format);
		my $terme = $match ? $termeLex{$match} : undef;
		if ( defined $terme ) {
			if ( $position{$terme} ) {
				if ( $exact )  {
					if ( $nr < $nf ) {
						$feuille->write_url($nr, 6, "$url_lex$position{$terme}", $format{'UrlFondVertTop1'}, $terme);
						}
					else	{
						$feuille->write_url($nr, 6, "$url_lex$position{$terme}", $format{'UrlFondVert'}, $terme);
						}
					}
				else	{
					if ( $nr < $nf ) {
						$feuille->write_url($nr, 6, "$url_lex$position{$terme}", $format{'UrlNormalTop1'}, $terme);
						}
					else	{
						$feuille->write_url($nr, 6, "$url_lex$position{$terme}", $format{'UrlNormal'}, $terme);
						}
					}
				}
			else	{
				if ( $nr < $nf ) {
					$feuille->write_string($nr, 6, $termeLex{$match}, $format{'Top1'});
					}
				else	{
					$feuille->write_string($nr, 6, $termeLex{$match});
					}
				}
			}
		}
	}

# Mise en place de la zone de validation des données
$debut = xl_rowcol_to_cell(1, 2);
$fin = xl_rowcol_to_cell($nr, 2);
# Liste des différents motifs de validation ou de rejet
my @libValid = (
	'Accepté (1 terme)', 
	'Accepté (2 termes)', 
	'Rejeté (syntaxe incorrecte)',
	'Rejeté (hors domaine)', 
	'Rejeté (non terminologique)',
	'Rejeté (trop générique)', 
	'Rejeté (autre) ⇒ commentaire'
	);
# Liste des différentes possibilités d'enrichissement
my @libEnrich = (
	'1 nouvelle entrée',
	'2 nouvelles entrées',
	'variante préférentiel existant', 
	'variante d’une nouvelle entrée',
	'remplacement d’un préférentiel',
	'déjà dans la ressource'
	);
$feuille->data_validation(
			"$debut:$fin",
			{
				validate => 'list',
				value    => [ @libValid ],
				input_title     => 'Choisissez une valeur',
				input_message   => 'dans la liste déroulante',
				error_title     => 'Erreur de saisie !',
				error_message   => 'Utilisez seulement les termes de la liste.',
			}
		);

$debut = xl_rowcol_to_cell(1, 8);
$fin = xl_rowcol_to_cell($nr, 8);
$feuille->data_validation(
			"$debut:$fin",
			{
				validate => 'list',
				value    => ['oui', 'non'],
#				input_title     => 'Choisissez une valeur',
#				input_message   => '« Oui » ou « Non »',
				error_title     => 'Erreur de saisie !',
				error_message   => 'Utilisez seulement « Oui » ou « Non ».',
			}
		);

$debut = xl_rowcol_to_cell(1, 10);
$fin = xl_rowcol_to_cell($nr, 10);
$feuille->data_validation(
			"$debut:$fin",
			{
				validate => 'list',
				value    => [ @libEnrich ],
				input_title     => 'Choisissez une valeur',
				input_message   => 'dans la liste déroulante',
				error_title     => 'Erreur de saisie !',
				error_message   => 'Utilisez seulement les termes de la liste.',
			}
		);

# Mise en place d'un format conditionnel colonne I
$debut = xl_rowcol_to_cell(1, 5);
$fin = xl_rowcol_to_cell($nr, 5);
$feuille->conditional_formatting(
			"$debut:$fin",
			{
			 type     => 'text',
			 criteria => 'containing',
			 value    => '☓',
			 format   => $format{'BoldCenterRed'},
			}
		);
$feuille->conditional_formatting(
			"$debut:$fin",
			{
			 type     => 'text',
			 criteria => 'containing',
			 value    => '⇒',
			 format   => $format{'BoldCenterGreen'},
			}
		);

# Mise en place d'un format conditionnel colonne I
$debut = xl_rowcol_to_cell(1, 9);
$fin = xl_rowcol_to_cell($nr, 9);
$feuille->conditional_formatting(
			"$debut:$fin",
			{
			 type     => 'text',
			 criteria => 'containing',
			 value    => '✕',	# Attention : pas la même croix !
			 format   => $format{'BoldCenterRed'},
			}
		);
$feuille->conditional_formatting(
			"$debut:$fin",
			{
			 type     => 'text',
			 criteria => 'containing',
			 value    => '☓',
			 format   => $format{'BoldCenterRed'},
			}
		);
$feuille->conditional_formatting(
			"$debut:$fin",
			{
			 type     => 'text',
			 criteria => 'containing',
			 value    => '⇒',
			 format   => $format{'BoldCenterGreen'},
			}
		);

# Mise en place du compteur de termes validés/refusés 
# sur la feuille de calcul "paramètres"
$feuille = $param;

$feuille->write(++ $np, 1, "Terme candidat :", $format{'Bold12'});
$feuille->write(++ $np, 2, "Validation", $format{'Bold12'});
$feuille->write($np, 6, "Récapitulatif", $format{'Bold12'});
$feuille->write($np, 10, "Alignement", $format{'Bold12'});
$feuille->write($np, 14, "Enrichissement", $format{'Bold12'});
# Validation du terme candidat

# Définition de quelques variables pour aligner les stats
my $nbase = $np;	# Ligne de base qui servira de référence
my $ntot  = $np + 4;	# Ligne où on écrira le total pour chaque colonne
if ( $np + $#libValid + 3 > $ntot ) {
	$ntot = $np + $#libValid + 3;
	}
if ( $np + $#libEnrich + 3 > $ntot ) {
	$ntot = $np + $#libEnrich + 3;
	}

$debut = xl_rowcol_to_cell(1, 2);
$fin = xl_rowcol_to_cell($nr, 2);

for ( my $nb = 0 ; $nb <= $#libValid ; $nb ++ ) {
	$feuille->write(++ $np, 2, $libValid[$nb], $format{'Left'});
	$feuille->write_formula($np, 3, "=COUNTIF(Résultat!$debut:$fin, \"$libValid[$nb]\")");
	$cell = xl_rowcol_to_cell($np, 3);
	$feuille->write_formula($np, 4, "=$cell/C6", $format{'Percent1d'});
	}

# Récapitulatif
$np = $nbase;
$feuille->write(++$np, 6, "Accepté", $format{'Left'});
my $nb = grep(/Accepté/, @libValid);
$debut = xl_rowcol_to_cell($nbase + 1, 3);
$fin = xl_rowcol_to_cell($nbase + $nb, 3);
$feuille->write_formula($np, 7, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($np, 7);
$feuille->write_formula($np, 8, "=$cell/C6", $format{'Percent1d'});
$feuille->write(++$np, 6, "Rejeté", $format{'Left'});
$debut = xl_rowcol_to_cell($nbase + $nb + 1, 3);
$fin = xl_rowcol_to_cell($nbase + $#libValid + 1, 3);
$feuille->write_formula($np, 7, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($np, 7);
$feuille->write_formula($np, 8, "=$cell/C6", $format{'Percent1d'});

# Validation de l'alignement
$debut = xl_rowcol_to_cell(1, 9);
$fin = xl_rowcol_to_cell($nr, 9);
$np = $nbase;
$feuille->write(++$np, 10, "Oui", $format{'Center'});
$feuille->write_formula($np, 11, "=COUNTIF(Résultat!$debut:$fin, \"☓\")");
$cell = xl_rowcol_to_cell($np, 11);
$feuille->write_formula($np, 12, "=$cell/C6", $format{'Percent1d'});
$feuille->write(++$np, 10, "Non", $format{'Center'});
$feuille->write_formula($np, 11, "=COUNTIF(Résultat!$debut:$fin, \"⇒\")");
$cell = xl_rowcol_to_cell($np, 11);
$feuille->write_formula($np, 12, "=$cell/C6", $format{'Percent1d'});


# Enrichissement
$debut = xl_rowcol_to_cell(1, 13);
$fin = xl_rowcol_to_cell($nr, 13);
$np = $nbase;
for ( my $nb = 0 ; $nb <= $#libEnrich ; $nb ++ ) {
	$feuille->write(++$np, 14, $libEnrich[$nb], $format{'Left'});
	$feuille->write_formula($np, 15, "=COUNTIF(Résultat!$debut:$fin, \"$libEnrich[$nb]\")");
	$cell = xl_rowcol_to_cell($np, 15);
	$feuille->write_formula($np, 16, "=$cell/C6", $format{'Percent1d'});
	}

# Écriture dse différents totaux
$debut = xl_rowcol_to_cell($nbase + 1, 3);
$fin = xl_rowcol_to_cell($nbase + $#libValid + 1, 3);

$feuille->write($ntot, 2, "Total", $format{'LeftItalic'});
$debut = xl_rowcol_to_cell($nbase + 1, 3);
$fin = xl_rowcol_to_cell($nbase + $#libValid + 1, 3);
$feuille->write_formula($ntot, 3, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($ntot, 3);
$feuille->write_formula($ntot, 4, "=$cell/C6", $format{'Percent1d'});

$debut = xl_rowcol_to_cell($nbase + 1, 7);
$fin = xl_rowcol_to_cell($nbase + 2, 7);

$feuille->write($ntot, 6, "Total", $format{'LeftItalic'});
$debut = xl_rowcol_to_cell($nbase + 1, 7);
$fin = xl_rowcol_to_cell($nbase + 2, 7);
$feuille->write_formula($ntot, 7, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($ntot, 7);
$feuille->write_formula($ntot, 8, "=$cell/C6", $format{'Percent1d'});

$debut = xl_rowcol_to_cell($nbase + 1, 11);
$fin = xl_rowcol_to_cell($nbase + 2, 11);

$feuille->write($ntot, 10, "Total", $format{'CenterItalic'});
$debut = xl_rowcol_to_cell($nbase + 1, 11);
$fin = xl_rowcol_to_cell($nbase + 2, 11);
$feuille->write_formula($ntot, 11, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($ntot, 11);
$feuille->write_formula($ntot, 12, "=$cell/C6", $format{'Percent1d'});

$debut = xl_rowcol_to_cell($nbase + 1, 15);
$fin = xl_rowcol_to_cell($nbase + $#libEnrich + 1, 15);

$feuille->write($ntot, 14, "Total", $format{'LeftItalic'});
$debut = xl_rowcol_to_cell($nbase + 1, 15);
$fin = xl_rowcol_to_cell($nbase + $#libEnrich + 1, 15);
$feuille->write_formula($ntot, 15, "=SUM($debut:$fin)");
$cell = xl_rowcol_to_cell($ntot, 15);
$feuille->write_formula($ntot, 16, "=$cell/C6", $format{'Percent1d'});

# Nombre de commentaires
$np = $ntot + 2;
$debut = xl_rowcol_to_cell(1, 3);
$fin = xl_rowcol_to_cell($nr, 3);
$feuille->write($np, 1, "Nombre de commentaires :", $format{'Bold12'});
$feuille->write_formula(++ $np, 3, "=COUNTA(Résultat!$debut:$fin)");
$cell = xl_rowcol_to_cell($np, 3);
$feuille->write_formula($np, 4, "=$cell/C6", $format{'Percent1d'});
$np ++;

# Affichage des messages d'erreurs 
if ( @messages ) {
	my $label = $#messages < 1 ? "Message d'erreur :" : "Messages d'erreur :";
	$feuille->write(++ $np, 1, $label, $format{'Bold12'});
	foreach my $message (@messages) {
		$feuille->write(++ $np, 2, $message);
		}
	}

# Dernière partie :
#  - ajout de la terminologie
my @intituLex = (
		'Termes génériques',
		'Terme préférentiel',
		'Termes spécifiques',
		'Synonymes',
		'Termes anglais',
		'Termes associés',
		'Note d’application',
		);

if ( $nbPages > 1 ) {
	$termino = substr($termino, 0, 31 - length($nbPages)) . "1";
	}

while ( $nbPages -- ) {
	# Ajout d'une feuille de calcul : la terminologie
	my $ressource = $excel->add_worksheet("$termino");
	$feuille = $ressource;

	$feuille->set_column(0, 5, 32);
	$feuille->set_column(6, 6, 40);

	$nc = 0;	# Numéro de colonne courante
	$nr = 0;	# Numéro de ligne courante sur la feuille "résultats"

	$format = $format{'BoldWrapCenterBrunPale14'};
	foreach my $intitule (@intituLex) {
		$feuille->write_string($nr, $nc ++, $intitule, $format);
		}

	# Maintien de l'entête (première rangée) 
	$feuille->freeze_panes(1, 0);

	# Remplissage de la feuille en cours
	my %dejavu = ();
	foreach my $terme (@termes) {
		next if $dejavu{$terme};
		next if $position{$terme} !~ /^$termino\!/;
		$nr ++;
		$dejavu{$terme} ++;
		$cell = xl_rowcol_to_cell($nr, 1);
		print LOG "\"$terme\" [$nbCells{$terme}] : $termino!$cell - $position{$terme}\n";
		if ( $position{$terme} ne "$termino!$cell" ) {
			die "\"$terme\" : $termino!$cell - $position{$terme}\n";
			}
		my $nbCells = $nbCells{$terme};
		my $dr = $nr + $nbCells - 1;	# Dernière Rangée
		if ( $dr > $nr ) {
			# Avec des exposants et des indices, utiliser "merge_range_type('rich_string', ...)"
			# Penser à regarder "rich_strings.pl" dans la distribution
			if ( $terme =~ /[\x{E000}-\x{E003}]/o ) {
				my @fragments = decoupe('Bold', 'BoldSup', 'BoldSub', $terme);
				$feuille->merge_range_type('rich_string', $nr, 1, $dr, 1, @fragments, $format{'Top1'});
				}
			else	{
				$feuille->merge_range_type('string', $nr, 1, $dr, 1, $terme, $format{'BoldTop1'});
				}
			my $tmp = join("\n", dedoublonne(@{$english{$terme}}));
			if ( $tmp ) {
				if ( $tmp =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $tmp);
					$feuille->merge_range_type('rich_string', $nr, 4, $dr, 4, @fragments, $format{'WrapTop1'});
					}
				else	{
					$feuille->merge_range_type('string', $nr, 4, $dr, 4, $tmp, $format{'WrapTop1'});
					}
				}
			else	{
				$feuille->merge_range($nr, 4, $dr, 4, undef, $format{'Top1'});
				}
			$tmp = join("\n", dedoublonne(@{$synonyme{$terme}}));
			if ( $tmp ) {
				if ( $tmp =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $tmp);
					$feuille->merge_range_type('rich_string', $nr, 3, $dr, 3, @fragments, $format{'WrapTop1'});
					}
				else	{
					$feuille->merge_range_type('string', $nr, 3, $dr, 3, $tmp, $format{'WrapTop1'});
					}
				}
			else	{
				$feuille->merge_range($nr, 3, $dr, 3, undef, $format{'Top1'});
				}
			if ( $note{$terme} ) {
				if ( $note{$terme} =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $note{$terme});
					$feuille->merge_range_type('string', $nr, 6, $dr, 6, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->merge_range_type('string', $nr, 6, $dr, 6, $note{$terme}, $format{'Top1'});
					}
				}
			else	{
				$feuille->merge_range($nr, 6, $dr, 6, undef, $format{'Top1'});
				}
			# Écriture du ou des génériques
			if ( @{$generique{$terme}} ) {
				$format = $format{'UrlInterneBleuTop1'};
				my $cell_format = $format{'Top1'};
				$base = $nr;
				foreach my $generique (dedoublonne(@{$generique{$terme}})) {
					if ( $generique =~ /[\x{E000}-\x{E003}]/o ) {
						my @fragments = decoupe('UrlInterneBleu', 'BleuSup', 'BleuSub', $generique);
						$feuille->write_url($base, 0, "internal:$position{$generique}", $format);
						$feuille->write_rich_string($base++, 0, @fragments, $cell_format);
						}
					else	{
						$feuille->write_url($base++, 0, "internal:$position{$generique}", $format, $generique);
						}
					$format = $format{'UrlInterneBleu'};
					$cell_format = undef;
					}
				}
			else	{
				$feuille->write($nr, 0, undef, $format{'Top1'});
				}
			# Écriture du ou des spécifiques
			if ( @{$specifique{$terme}} ) {
				$format = $format{'UrlInterneBrunTop1'};
				my $cell_format = $format{'Top1'};
				$base = $nr;
				foreach my $specifique (dedoublonne(@{$specifique{$terme}})) {
					if ( $specifique =~ /[\x{E000}-\x{E003}]/o ) {
						my @fragments = decoupe('UrlInterneBrun', 'BrunSup', 'BrunSub', $specifique);
						$feuille->write_url($base, 2, "internal:$position{$specifique}", $format);
						$feuille->write_rich_string($base++, 2, @fragments, $cell_format);
						}
					else	{
						$feuille->write_url($base++, 2, "internal:$position{$specifique}", $format, $specifique);
						}
					$format = $format{'UrlInterneBrun'};
					$cell_format = undef;
					}
				}
			else	{
				$feuille->write($nr, 2, undef, $format{'Top1'});
				}
			# Écriture du ou des termes associés
			if ( @{$associe{$terme}} ) {
				$format = $format{'UrlInterneVertTop1'};
				my $cell_format = $format{'Top1'};
				$base = $nr;
				foreach my $associe (dedoublonne(@{$associe{$terme}})) {
					if ( $associe =~ /[\x{E000}-\x{E003}]/o ) {
						my @fragments = decoupe('UrlInterneVert', 'VertSup', 'VertSub', $associe);
						$feuille->write_url($base, 5, "internal:$position{$associe}", $format);
						$feuille->write_rich_string($base++, 5, @fragments, $cell_format);
						}
					else	{
						$feuille->write_url($base++, 5, "internal:$position{$associe}", $format, $associe);
						}
					$format = $format{'UrlInterneVert'};
					$cell_format = undef;
					}
				}
			else	{
				$feuille->write($nr, 5, undef, $format{'Top1'});
				}
			$nr = $dr;
			}
		else	{
			# Avec des exposants et des indices, utiliser "write_rich_string(...)"
			if ( $terme =~ /[\x{E000}-\x{E003}]/o ) {
				my @fragments = decoupe('Bold', 'BoldSup', 'BoldSub', $terme);
				$feuille->write_rich_string($nr, 1, @fragments, $format{'Top1'});
				}
			else	{
				$feuille->write_string($nr, 1, $terme, $format{'BoldTop1'});
				}
			my $tmp = ${$english{$terme}}[0];
			if ( $tmp ) {
				if ( $tmp =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $tmp);
					$feuille->write_rich_string($nr, 4, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_string($nr, 4, $tmp, $format{'Top1'});
					}
				}
			else	{
				$feuille->write($nr, 4, undef, $format{'Top1'});
				}
			$tmp = ${$synonyme{$terme}}[0];
			if ( $tmp ) {
				if ( $tmp =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $tmp);
					$feuille->write_rich_string($nr, 3, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_string($nr, 3, $tmp, $format{'Top1'});
					}
				}
			else	{
				$feuille->write($nr, 3, undef, $format{'Top1'});
				}
			if ( $note{$terme} ) {
				if ( $note{$terme} =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('', 'Sup', 'Sub', $note{$terme});
					$feuille->write_rich_string($nr, 6, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_string($nr, 6, $note{$terme}, $format{'Top1'});
					}
				}
			else	{
				$feuille->write($nr, 6, undef, $format{'Top1'});
				}
			# Écriture du générique (s'il existe)
			if ( @{$generique{$terme}} ) {
				my $generique = ${$generique{$terme}}[0];
				if ( $generique =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('UrlInterneBleu', 'BleuSup', 'BleuSub', $generique);
					$feuille->write_url($nr, 0, "internal:$position{$generique}", $format{'UrlInterneBleuTop1'});
					$feuille->write_rich_string($nr, 0, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_url($nr, 0, "internal:$position{$generique}", $format{'UrlInterneBleuTop1'}, $generique);
					}
				}
			else	{
				$feuille->write($nr, 0, undef, $format{'Top1'});
				}
			# Écriture du spécifique (s'il existe)
			if ( @{$specifique{$terme}} ) {
				my $specifique = ${$specifique{$terme}}[0];
				if ( $specifique =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('UrlInterneBrun', 'BrunSup', 'BrunSub', $specifique);
					$feuille->write_url($nr, 2, "internal:$position{$specifique}", $format{'UrlInterneBrunTop1'});
					$feuille->write_rich_string($nr, 2, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_url($nr, 2, "internal:$position{$specifique}", $format{'UrlInterneBrunTop1'}, $specifique);
					}
				}
			else	{
				$feuille->write($nr, 2, undef, $format{'Top1'});
				}
			# Écriture du terme associé (s'il existe)
			if ( @{$associe{$terme}} ) {
				my $associe = ${$associe{$terme}}[0];
				if ( $associe =~ /[\x{E000}-\x{E003}]/o ) {
					my @fragments = decoupe('UrlInterneVert', 'VertSup', 'VertSub', $associe);
					$feuille->write_url($nr, 5, "internal:$position{$associe}", $format{'UrlInterneVertTop1'});
					$feuille->write_rich_string($nr, 5, @fragments, $format{'Top1'});
					}
				else	{
					$feuille->write_url($nr, 5, "internal:$position{$associe}", $format{'UrlInterneVertTop1'}, $associe);
					}
				}
			else	{
				$feuille->write($nr, 5, undef, $format{'Top1'});
				}
			}
		}
	$termino ++;
	}

# Fermeture du fichier Excel
$excel->close();


exit 0;



sub usage
{
my $retour = shift;

print STDERR "Usage : $programme -c candidats -l lexique -e fichier Excel \n"; 
print STDERR " " x (length($programme) + 6), "   -u URL Totem ( -a Aroma | -s Smoa )\n";
print STDERR " " x (length($programme) + 6), " [ -r rapport ] [ -d domaine ] \n";
print STDERR " " x (length($programme) + 6), " [ -v 'nom du vocabulaire' ] \n";
# print STDERR "Usage : $programme -c candidats -l lexique [ -l autre lexique ]* \n"; 
# print STDERR " " x (length($programme) + 6), "   -e fichier Excel ( -a Aroma | -s Smoa )\n";
# print STDERR " " x (length($programme) + 6), " [ -r rapport ] [ -d domaine ] \n";
# print STDERR " " x (length($programme) + 6), " [ -v 'nom du vocabulaire' ] \n";

exit $retour;
}

sub traiteTBX
{
my $lexique = shift;

# Initialisation du parseur et ...
my $twig = XML::Twig->new( 
			twig_roots => {
				'termEntry' => 1,
				},
			twig_handlers => {
				'termEntry' => \&termLex,
				'descrip'   => \&descripLex,
				'langSet'   => \&langSet,
				},
			);

# ... parsage du fichier
$twig->parsefile("$lexique");
$twig->purge;

# Complétion des information dans les différents hachages
foreach my $terme (keys %generique) {
	foreach my $item (@{$generique{$terme}}) {
		if ( $termeLex{$item} ) {
			$item = $termeLex{$item};
			}
		else	{
			print LOG "Pas de terme pour l'identifiant \"$item\" [Générique de \"$terme\"]\n";
			$item = "";
			}
		}
	}

foreach my $id (keys %specifique) {
	if ( $termeLex{$id} ) {
		push(@{$specifique{$termeLex{$id}}}, @{$specifique{$id}});
		delete $specifique{$id};
		}
	else	{
		print LOG "Pas de terme pour l'identifiant \"$id\" [générique du spécifique \"${$specifique{$id}}[0]\"]\n";
		}
	}

foreach my $terme (keys %associe) {
	foreach my $item (@{$associe{$terme}}) {
		if ( $termeLex{$item} ) {
			$item = $termeLex{$item};
			}
		else	{
			print LOG "Pas de terme pour l'identifiant \"$item\" [Associé à \"$terme\"]\n";
			$item = "";
			}
		}
	}
}

sub traiteRDF
{
my $lexique = shift;

# Initialisation du parseur et ...
my $twig = XML::Twig->new( 
			twig_roots => {
				'rdf:Description' => 1,
				},
			twig_handlers => {
				'rdf:Description'               => \&description,
				'rdfs:label[@xml:lang="fr"]'    => sub {$termeLex = $_->text;},
				'rdfs:label[@xml:lang="en"]'    => sub {@english = ($_->text);},
				'altLabel[@xml:lang="fr"]'      => sub {push(@synonymes, $_->text);},
				'skos:altLabel[@xml:lang="fr"]' => sub {push(@synonymes, $_->text);},
				'altLabel[@xml:lang="en"]'      => sub {push(@english, $_->text);},
				'skos:altLabel[@xml:lang="en"]' => sub {push(@english, $_->text);},
				'scopeNote'                     => sub {$note = $_->text;},
				},
			);

# ... parsage du fichier
$twig->parsefile("$lexique");
$twig->purge;

# Complétion des information dans les différents hachages
foreach my $id (keys %generique) {
	if ( not $termeLex{$id} ) {
		print LOG "Pas de terme pour l'identifiant \"$id\" [spécifique]\n";
		next;
		}
	my $terme = $termeLex{$id};
	my @tmp = ();
	foreach my $item (@{$generique{$id}}) {
		if ( $termeLex{$item} ) {
			push(@tmp, $termeLex{$item});
			}
		else	{
			print LOG "Pas de terme pour l'identifiant \"$item\" [Générique de \"$terme\"]\n";
			}
		}
	delete $generique{$id};
	push(@{$generique{$terme}}, @tmp);
	}

foreach my $id (keys %specifique) {
	if ( not $termeLex{$id} ) {
		print LOG "Pas de terme pour l'identifiant \"$id\" [générique]\n";
		next;
		}
	my $terme = $termeLex{$id};
	my @tmp = ();
	foreach my $item (@{$specifique{$id}}) {
		if ( $termeLex{$item} ) {
			push(@tmp, $termeLex{$item});
			}
		else	{
			print LOG "Pas de terme pour l'identifiant \"$item\" [Spécifique de \"$terme\"]\n";
			}
		}
	delete $specifique{$id};
	push(@{$specifique{$terme}}, @tmp);
	}

foreach my $id (keys %associe) {
	if ( not $termeLex{$id} ) {
		print LOG "Pas de terme pour l'identifiant \"$id\" [Associé]\n";
		next;
		}
	my $terme = $termeLex{$id};
	my @tmp = ();
	foreach my $item (@{$associe{$id}}) {
		if ( $termeLex{$item} ) {
			push(@tmp, $termeLex{$item});
			}
		else	{
			print LOG "Pas de terme pour l'identifiant \"$item\" [Associé à \"$terme\"]\n";
			}
		}
	delete $associe{$id};
	push(@{$associe{$terme}}, @tmp);
	}
}

sub termEntry
{
my ($twig, $element) = @_;

my $id = $element->att('xml:id');
$candidats{$id}{'lemme'} = $lemme;
if ( $id{$lemme} ) {
	print STDERR "Double Id pour lemme \"$lemme\" : $id{$lemme} et $id\n";
#	exit 5;
	}
$id{$lemme} = $id;
$candidats{$id}{'pilote'} = $pilote;
$candidats{$id}{'pos'} = $pos;
if ( $pilote ne $lemme ) {
	if ( $id{$pilote} ) {
		print STDERR "Double Id pour pilote \"$pilote\" : $id{$pilote} et $id\n";
		my $message = sprintf("Terme pilote \"%s\" en double (lemmes \"%s\" et \"%s\")", 
					$pilote, $candidats{$id{$pilote}}{'lemme'}, $lemme);
		push(@messages, $message);
		}
	else	{
		$id{$pilote} = $id;
		}
	}
foreach my $item (keys %synonymes) {
	next if $item eq $lemme or $item eq $pilote;
	$candidats{$id}{'forme'}{$item} ++;
	if ( $id{$item} ) {
		print STDERR "Double Id pour forme \"$item\" : $id{$item} et $id\n";
		my $message = sprintf("Terme fléchi \"%s\" en double (lemmes \"%s\" et \"%s\")", 
				      $item, $candidats{$id{$item}}{'lemme'}, $lemme);
		push(@messages, $message);
		}
	else	{
		$id{$item} = $id;
		}
	}

$lemme = $pilote = $pos = "";
%synonymes = ();

$element->purge();
}

sub descrip
{
my ($twig, $element) = @_;

# my $type = $element->att('type');

my $formes = $element->text;
$formes =~ s/\[\{(.+)\}\]/$1/o;
my @formes = split(/\}, \{/, $formes);
foreach my $forme (@formes) {
	my ($terme) = $forme =~ /term="(.+?)", /o;
	$synonymes{$terme} ++
	}
}


sub termLex
{
my ($twig, $element) = @_;

if ( $termeLex ) {
	push(@termes, $termeLex);
	}
else	{
	return;
	}

my $id = $element->att('xml:id');
$id = substr($id, 3) if $id =~ /^BV\./o;

$termeLex{$id} = $termeLex;

if ( @english ) {
	push(@{$english{$termeLex}}, @english);
	@english = ();
	}
else	{
	print LOG "Pas de traduction anglaise pour \"$termeLex\" [$id]\n";
	}

if ( @synonymes ) {
	foreach my $synonyme (@synonymes) {
		push(@{$synonyme{$termeLex}}, $synonyme);
		$synonLex{$id}{$synonyme} ++;
		}
	@synonymes = ();
	}

if ( @generiques ) {
	while ( my $generique = shift @generiques ) {
		push(@{$specifique{$generique}}, $termeLex);
		push(@{$generique{$termeLex}}, $generique);
		}
	}

if ( @associes ) {
	while ( my $associe = shift @associes ) {
		push(@{$associe{$termeLex}}, $associe);
		}
	}

if ( $note ) {
	$note{$termeLex} = $note;
	}

$termeLex = $langue = $note = "";

$element->purge;
}

sub langSet
{
my ($twig, $element) = @_;

$langue = $element->att('xml:lang');

foreach my $child ( $element->children('tig') ) {
	tig($twig, $child);
	}
}

sub tig
{
my ($twig, $element) = @_;

my $mot = $element->first_child_text('tei:term');

my $statut = "";
foreach my $child ( $element->children('termNote') ) {
	my $type = $child->att('type');
	if ( $type eq 'administrativeStatus' ) {
		$statut = $child->text;
		}
	}

if ( $langue eq 'fr' ) {
	if ( $statut eq 'preferredTerm' ) {
		$termeLex = $mot;
		}
	else	{
		push(@synonymes, $mot);
		}
	}
elsif ( $langue eq 'en' )	{
	push(@english, $mot);
	}
}

sub descripLex
{
my ($twig, $element) = @_;

my $type = $element->att('type');

if ( $type eq 'broaderConceptGeneric' ) {
	my $numero = $element->text;
	$numero = substr($numero, 3) if $numero =~ /^BV\./o;
	push(@generiques, $numero)
	}
elsif ( $type eq 'relatedConcept' ) {
	my $numero = $element->text;
	$numero = substr($numero, 3) if $numero =~ /^BV\./o;
	push(@associes, $numero);
	}
elsif ( $type eq 'explanation' ) {
	$note = $element->text;
	}
}

sub description
{
my ($twig, $element) = @_;

my $description = $element->att('rdf:about');
if ( $description =~ m|/(\d+)\z| ) {
	my $id = $1;
	if ( $termeLex ) {
		$termeLex{$id} = $termeLex;
		push(@termes, $termeLex);
		}
	else	{
		print LOG "Pas de terme pour l'identifiant \"$description\"\n";
		return;
		}

	if ( @english ) {
		push(@{$english{$termeLex}}, @english);
		@english = ();
		}
	else	{
		print LOG "Pas de traduction anglaise pour \"$termeLex\" [$id]\n";
		}

	if ( @synonymes ) {
		push(@{$synonyme{$termeLex}}, @synonymes);
		@synonymes = ();
		}

	if ( $note ) {
		$note{$termeLex} = $note;
		$note = "";
		}

	$termeLex = "";
	}

elsif ( $description =~ /^assoc:_\d+\z/o ) {
	my $bt = $element->first_child('bt');
	my $nt = $element->first_child('nt');
	my @rt = $element->children('rt');

	if ( $bt and $nt ) {
		(my $id1 = $bt->att('rdf:resource')) =~ s|^.*/(\d+)\z|$1|;
		(my $id2 = $nt->att('rdf:resource')) =~ s|^.*/(\d+)\z|$1|;
		push(@{$generique{$id2}}, $id1);
		push(@{$specifique{$id1}}, $id2);
		}
	elsif ( @rt ) {
		(my $id1 = $rt[0]->att('rdf:resource')) =~ s|^.*/(\d+)\z|$1|;
		(my $id2 = $rt[1]->att('rdf:resource')) =~ s|^.*/(\d+)\z|$1|;
		push(@{$associe{$id1}}, $id2);
		push(@{$associe{$id2}}, $id1);
		}
	}

$element->purge;
}

sub traite
{
my $alignement = shift;

# Initialisation du parseur et ...
my $twig = XML::Twig->new( 
			   twig_roots => {
				'Cell' => 1,
				},
			   twig_handlers => {
#				'entity1/edoal:Class' => sub {$ent1 = $_->att('rdf:about')},
#				'entity2/edoal:Class' => sub {$ent2 = $_->att('rdf:about')},
				'entity1'             => sub {$ent1 = $_->att('rdf:resource')},
				'entity2'             => sub {$ent2 = $_->att('rdf:resource')},
				'relation'            => sub {$relation = $_->text;},
				'measure'             => sub {$score = $_->text;},
				'Cell'                => \&cell,
				},
			  );


# ... parsage du fichier
$twig->parsefile("$alignement");
$twig->purge;
}

sub cell
{
my ($twig, $element) = @_;

$ent1 =~ s|(.*/)||o;
$ent2 =~ s|(.*/)||o;
$relation =~ s/.*\.(\w+)Relation/$1/o;
$align->{$ent1}{$ent2} = "$relation:$score";

$score = 0.0;
$ent1 = $ent2 = $relation = "";
$element->purge;
}


sub compare
{
my ($id, @liste) = @_;

my %liste = (lc($termeLex{$id}) => 1);
foreach my $item (sort keys %{$synonLex{$id}}) {
	$liste{lc($item)} ++;
	}

foreach my $terme (@liste) {
	return 1 if $liste{lc($terme)};
	}

return 0;
}

sub tri
{
my @lignes = @_;

my %valeur  = ();
my %methode = ();
my %match   = ();
foreach my $ligne (@lignes) {
	my ($score, $methode, $match) = $ligne =~ /^(\d?\.\d*)\t(\w+)\t.+?\t(.+?)\t/o;
	$valeur{$ligne} = $score;
	if ( $methode eq "Aroma" ) {
		$methode{$ligne} = 2;
		}
	elsif ( $methode eq "Smoa" ) {
		$methode{$ligne} = 1;
		}
	else	{
		$methode{$ligne} = 0;
		}
	$match{$ligne} = $match;
	}

my @tmp = ();
my %tmp = ();
foreach my $ligne (sort {$valeur{$b} <=> $valeur{$a} or $methode{$b} <=> $methode{$a}} @lignes) {
	next if $tmp{$match{$ligne}};
	$tmp{$match{$ligne}} ++;
	push(@tmp, $ligne);
	}

return @tmp;
}

sub dedoublonne
{
my @liste  = @_;

my @sortie = ();
my %dejavu = ();

foreach my $item (@liste) {
	next if $dejavu{$item};
	push(@sortie, $item);
	$dejavu{$item} ++;
	}

return @sortie;
}