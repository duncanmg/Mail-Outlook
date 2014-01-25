#!/usr/bin/perl -w
use strict;

# no strict 'subs';

use Test::More tests => 14;

use lib 't/testlib';

my $tests = 14;

eval {

  SKIP: {
        eval "use Typelibs";
        skip "Microsoft Outlook doesn't appear to be installed\n", $tests
          if ($@);

        my $vers = Typelibs::ExistsTypeLib('Microsoft Outlook');
        skip "Microsoft Outlook doesn't appear to be installed\n", $tests
          unless ($vers);

        eval "use Mail::Outlook";
        skip "Unable to make a connection to Microsoft Outlook\n", $tests
          if ($@);

        my $outlook = Mail::Outlook->new('Inbox');
        my $folder  = $outlook->folder();
        isa_ok( $folder, 'Mail::Outlook::Folder' );

        my $message = $folder->first;
        isa_ok( $message, 'Mail::Outlook::Message' );

        my $name = $message->From();
        skip "Access to Microsoft Outlook has been declined", ( $tests - 2 )
          unless ($name);
        ok( $name, "name is true" );

        $message = $folder->next;
        isa_ok( $message, 'Mail::Outlook::Message' );
        ok( $message->From(),
            "folder->next() return a message with a true From method." );

        $folder  = $outlook->folder('Inbox');
        $message = $folder->last;
        isa_ok( $message, 'Mail::Outlook::Message' );
        ok( $message->From(),
            "folder->last() return a message with a true From method." );

        $message = $folder->previous;
        isa_ok( $message, 'Mail::Outlook::Message' );
        ok( $message->From(),
            "folder->previous() return a message with a true From method." );

        $folder = $outlook->folder('Inbox/ANameThatShouldNotExist');
        is( $folder, undef,
            'Got undef when looking for Inbox/ANameThatShouldNotExist' );
        $folder = $outlook->folder('ANameThatShouldNotExist');
        is( $folder, undef,
            'Got undef when looking for ANameThatShouldNotExist' );

        eval {
            my $olFolderInbox = 0;

            eval "use Win32::OLE::Constttt 'Microsoft Outlook'; "
              . "\$olFolderInbox = olFolderInbox;";
            ok( !$olFolderInbox, "Correctly handling non-existent package." );
            eval "use Win32::OLE::Const 'Microsoft Outlookkkkk'; "
              . "\$olFolderInbox = olFolderInbox;";
            ok( !$olFolderInbox, "Correctly handling non-existent library." );

            eval "use Win32::OLE::Const 'Microsoft Outlook'; "
              . "\$olFolderInbox = olFolderInbox;";

          SKIP: {
                skip "Unable to make a connection to Microsoft Outlook. $@", 1
                  if !$olFolderInbox;
                $folder = $outlook->folder($olFolderInbox);
                isa_ok( $folder, 'Mail::Outlook::Folder' );

                1;
            }
        } || diag($@);

    }

};

if ( $@ =~ /Network problems/ ) {
    skip "Microsoft Outlook cannot connect to the server.\n", $tests;
    exit;
}
