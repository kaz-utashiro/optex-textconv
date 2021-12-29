package App::optex::textconv::ooxml;

our $VERSION = '0.1401';

use v5.14;
use warnings;
use Carp;

use App::optex::textconv::Converter 'import';

our @CONVERTER;

use App::optex::textconv::ooxml::regex;

eval {
    require App::optex::textconv::ooxml::xslt;
} and do {
    import  App::optex::textconv::ooxml::xslt;
};

1;
