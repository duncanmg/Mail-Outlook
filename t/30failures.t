#!/usr/bin/perl -w
use strict;

# DO NOT COPY THESE TWO LINES UNLESS YOU UNDERSTAND WHAT THEY DO.
# ... AND EVEN THEN DON'T COPY THEM!
use strict;
*strict::import = sub { $^H };

use Test::More tests => 3;

my $nomock;
my $mock;

BEGIN {
	eval "use Test::MockObject";
    $nomock = $@;

    unless($nomock) {
        $mock = Test::MockObject->new();
        $mock->fake_module( 'Win32::OLE' );
        $mock->fake_new( 'Win32::OLE' );
    }
}

use lib qw(./t/fake);

SKIP: {
	my $tests = 3;

	skip "Test::MockObject doesn't appear to be installed\n", $tests	if($nomock);

	eval "use Mail::Outlook";
	skip "Unable to make a fake connection to Microsoft Outlook: $@\n",$tests	if($@);

    use Win32::OLE::Const   'Microsoft::Outlook';

    $mock->mock( 'GetNameSpace', sub { return undef } );
	$mock->mock( 'GetActiveObject', sub { die "Forced Failure" } );
	my $outlook = Mail::Outlook->new();
	is($outlook,undef);

	$mock->mock( 'GetActiveObject', sub { return undef } );
	$outlook = Mail::Outlook->new();
	is($outlook,undef);

	$outlook = Mail::Outlook->new();
	is($outlook,undef);
}
