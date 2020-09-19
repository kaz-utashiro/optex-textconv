package App::optex::textconv::msdoc;

our $VERSION = '0.08';

use v5.14;
use warnings;
use Carp;
use utf8;

use App::optex::textconv::Converter 'import';

our @CONVERTER = (
    [ qr/\.docx$/ => \&to_text ],
    [ qr/\.pptx$/ => \&to_text ],
    [ qr/\.xlsx$/ => \&to_text ],
    );

sub extract_text {
    local $_ = shift;
    my $type = shift;
    my $xml_re = qr/<\?xml\b[^>]*\?>\s*/;
    return $_ unless /$xml_re/;

    my @xml  = grep { length } split /$xml_re/;
    my @text = map  { _xml2text($_, $type) } @xml;
    join "\n", @text;
}

my %param = (
    docx => { space => 2, separator => ""   },
    xlsx => { space => 1, separator => "\t" },
    pptx => { space => 1, separator => ""   },
    );

my $replace_reference = do {
    my %hash = qw( amp &  lt <  gt > );
    my @keys = keys %hash;
    my $re = do { local $" = '|'; qr/&(@keys);/ };
    sub { s/$re/$hash{$1}/g }
};

sub _xml2text {
    local $_ = shift;
    my $type = shift;
    my $param = $param{$type} or die;

    my @p;
    while (m{<(?<tag>[apw]:p|si)\b[^>]*>(?<para>.*?)</\g{tag}>}sg) {
	my $p = $+{para};
	my @s;
	while ($p =~ m{
	       (?<tab> <w:tab/> | <w:tabs> )
	       |
	       <(?<tag>(?:[apw]:)?t)\b[^>]*> (?<text>[^<]*?) </\g{tag}>
	       }xsg) {
	    if ($+{tab}) {
		push @s, "  ";
	    } else {
		push @s, $+{text} if $+{text} ne '';
	    }
	}
	@s or next;
	push @p, join($param->{separator}, @s) . ("\n" x $param->{space});
    }
    my $text = join '', @p;
    $replace_reference->() for $text;
    $text;
}

use App::optex::textconv::Zip;

sub to_text {
    my $zipfile = shift;
    my $zip = App::optex::textconv::Zip->new($zipfile);
    my $type = $zip->suffix or return;
    join "\n", map {
	my $text = extract_text $zip->extract($_), $type;
	$text ne "" ? ("[ $_ ]\n", $text) : ();
    }
    $zip->list;
}

1;
