package App::optex::textconv::tika;

our $VERSION = '0.11';

use v5.14;
use warnings;
use Carp;

use App::optex::textconv::Converter 'import';

our @CONVERTER = (
    [ qr/\.docx$/ => \&to_text ],
    [ qr/\.pptx$/ => \&to_text ],
    [ qr/\.xlsx$/ => \&to_text ],
    );

sub to_text {
    my $file = shift;
    my $format = q(tika --text "%s");
    my $exec = sprintf $format, $file;
    qx($exec);
}

1;
