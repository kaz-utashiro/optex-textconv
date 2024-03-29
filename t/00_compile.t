use strict;
use Test::More 0.98;

use_ok $_ for qw(
    App::optex::textconv
    App::optex::tc
    App::optex::textconv::Converter
    App::optex::textconv::default
    App::optex::textconv::msdoc
    App::optex::textconv::pandoc
    App::optex::textconv::pdf
    App::optex::textconv::tika
    App::optex::textconv::ooxml
    App::optex::textconv::ooxml::regex
);

TODO: {
    local $TODO = 'May not be installed';
    use_ok $_ for qw(
	App::optex::textconv::ooxml::xslt
    );
}

done_testing;

