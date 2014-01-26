#!/usr/bin/perl -w
use strict;
use warnings;

use Test::More;
use Test::Output qw/combined_from/;

use_ok('Win32::OLE')                || goto END;
use_ok('Alien::Microsoft::Outlook') || goto END;
use_ok('Mail::Outlook')             || goto END;

my $outlook = Mail::Outlook->new('Inbox');
ok( $outlook,
    "Created a Mail::Outlook object using the constructor argument 'Inbox'" );

my $folder = $outlook->folder();
isa_ok( $folder, 'Mail::Outlook::Folder' );
ok( !Win32::OLE->LastError(), "No Win32::OLE::LastError" );

test_basic_operations($folder);

# Win32::OLE emits some noise which we aren't interested in unless the test fails.
my $combined_output = combined_from(
    sub {
        $folder = $outlook->folder('Inbox/ANameThatShouldNotExist');
    }
);
is( $folder, undef, 'Got undef when looking for Inbox/ANameThatShouldNotExist' )
  || diag($combined_output);

$folder = $outlook->folder('ANameThatShouldNotExist');
is( $folder, undef, 'Got undef when looking for ANameThatShouldNotExist' );

END:

done_testing;

# This won't work unless there are at least two messages in the folder.
sub test_basic_operations {
    my $folder = shift;

    test_operation( sub { $folder->first; }, "first" );

    test_operation( sub { $folder->next; }, "next" );

    test_operation( sub { $folder->last; }, "last" );

    test_operation( sub { $folder->previous; }, "previous" );

    return 1;

}

sub test_operation {
    my ( $command, $message ) = @_;
    my $mail_message = $command->();
    ok( $mail_message,
        "Retrieved a mail message using the method \"$message\"." );
    ok( !Win32::OLE->LastError(), "No Win32::OLE::LastError" );
    isa_ok( $mail_message, 'Mail::Outlook::Message' ) || return;
    ok( $mail_message->From(),
        "The message has a From method with returns true. "
          . $mail_message->From() );
    return 1;
}

