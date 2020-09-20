package App::optex::textconv::Zip;

use v5.14;
use warnings;

sub new {
    my $class = shift;
    my $zipfile = shift;
    bless {
	NAME => $zipfile,
	FILE => {},
    }, $class;
}

sub name {
    my $obj = shift;
    $obj->{NAME};
}

sub suffix {
    my $obj = shift;
    $obj->{SUFFIX} //= do {
	($obj->name =~ /\.(docx|xlsx|pptx)$/)[0] or "";
    };
}

sub list {
    my $obj = shift;
    my $list = $obj->{LIST} //= do {
	my $zipfile = $obj->name;
	my $suffix = $obj->suffix;
	my @entry = do {
	    if ($suffix eq 'docx') {
		my %select = map { $_ => 1 } qw(document endnotes footnotes);
		do {
		    map  { $_->[0] }
		    grep { $select{$_->[1]} }
		    map  { m{(word/(.*)\.xml)$} ? [ $1, $2 ] : () }
		    `unzip -l \"$zipfile\" word/*.xml`;
		};
	    }
	    elsif ($suffix eq 'xlsx') {
		map { "xl/$_.xml" } qw(sharedStrings);
	    }
	    elsif ($suffix eq 'pptx') {
		map  { $_->[0] }
		sort { $a->[1] <=> $b->[1] }
		map  { m{(ppt/slides/slide(\d+)\.xml)$} ? [ $1, $2 ] : () }
		`unzip -l \"$zipfile\" ppt/slides/slide*`;
	    }
	};
	\@entry;
    };
    @{$list};
}

sub extract {
    my $obj = shift;
    my $file = shift;
    $obj->{FILE}->{$file} //= do {
	my $zipfile = $obj->name;
	`unzip -p \"$zipfile\" \"$file\"`;
    };
}

1;
