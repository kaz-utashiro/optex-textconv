package App::optex::textconv::ooxml;

our $VERSION = '0.13';

use v5.14;
use warnings;
use Carp;

use App::optex::textconv::Converter 'import';

require App::optex::textconv::ooxml::regex;
require App::optex::textconv::ooxml::xslt;

our @CONVERTER = (
    [ qr/\.doc[xm]$/ => \&App::optex::textconv::ooxml::xslt::to_text ],
    [ qr/\.ppt[xm]$/ => \&App::optex::textconv::ooxml::xslt::to_text ],
    [ qr/\.xls[xm]$/ => \&App::optex::textconv::ooxml::regex::to_text ],
    );

1;
