#####################################################################
#
# Win32::ASP - a Module for ASP (PerlScript) Programming
#
# Author: Matt Sergeant
# Revision: 1.10
# Last Change: Added SetCookie (version bumbed to 1.10 for CPAN)
#####################################################################
# Copyright 1998 Matt Sergeant.  All rights reserved.
#
# This file is distributed under the Artistic License. See
# http://www.ActiveState.com/corporate/artistic_license.htm or
# the license that comes with your perl distribution.
#
# The most up to date version of this script should always be available
# at http://come.to/fastnet
#
# For comments, questions, bugs or general interest, feel free to
# contact me at msergeant@ndirect.co.uk
#####################################################################

package Win32::ASP;
use strict;

=head1 NAME

Win32::ASP - a Module for ASP (PerlScript) Programming

=head1 Synopsis

	use Win32::ASP;

	print "This is a test<BR><BR>";

	$PageName = GetFormValue('PageName');
	if($PageName eq 'Select a page...') {
		die "Please go back and select a value from the Pages list";
	}

	print "You selected the ", $PageName, " page";
	exit;

=head1 Description

These routines are some I knocked together one day when I was saying the
following: "Why don't my "print" statements output to the browser?" and
"Why doesn't exit and die end my script?". So I started investigating how
I could overload the core functions. "print" is overloaded via the tie
mechanism (thanks to Eryq (F<eryq@zeegee.com>), Zero G Inc. for the
code which I ripped from IO::Scalar). You can also get at print using the
OO mechanism with $Win32::ASP::SH->print(). Also added recently was code
that allowed cleanup stuff to be executed when you exit() or die(), this
comes in the form of the C<AddDeathHook> function. The C<BinaryWrite> function
simply wraps up unicode conversion and BinaryWrite in one call. Finally I
was annoyed that I couldn't just develop a script
using GET and then change to POST for release because of the difference in
how the ASP code handles the different formats, GetFormValue solves that one.

=head2 Installation instructions

Download the file. If you get it from my geocities web site you will have
to rename it (they don't like .pm extensions). Make sure the file is called
ASP.pm

Now move the file into the perl\site\<version>\lib\win32 directory (where
<version> is your perl version).

That's it - you're ready to start PerlScripting with ease.

=head1 Function Reference

=cut

BEGIN {
	use Exporter ();
	use vars		qw( @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS
						$Application
						$ObjectContext
						$Request
						$Response
						$Server
						$Session
						@DeathHooks
						);

	@ISA		=	qw(Exporter);
	@EXPORT		=	qw(	Print
						wprint
						die
						exit
						GetFormValue
						GetFormCount
					);
	%EXPORT_TAGS =	( 'strict' => [qw(
						Print
						wprint
						die
						exit
						GetFormValue
						GetFormCount
						$Application
						$ObjectContext
						$Request
						$Response
						$Server
						$Session
						)]
					);
	@EXPORT_OK = qw ( SetCookie );
	Exporter::export_ok_tags('strict'); # Add all strict vars to @EXPORT_OK

	$Application = $main::Application;
	$ObjectContext = $main::ObjectContext;
	$Request = $main::Request;
	$Response = $main::Response;
	$Server = $main::Server;
	$Session = $main::Session;

#	my $Servs = $Request->ServerVariables;
#	foreach my $env (in $Servs) {
#		$ENV{$env} = $Request->ServerVariables($env);
#	}

}

$Win32::ASP::VERSION='1.10';

my $SH = tie *RESPONSE_FH, 'Win32::ASP::IO';
select RESPONSE_FH;

=over 4

=item use Win32::ASP qw(:strict);

This allows you to use the ASP module in a "strict" perl script. Normally under "use strict"
PerlScript would complain that the ASP objects ($Response, $Session etc) were not
initialised: "Global symbol "Response" requires explicit package name at - line XXX".

To get around this I simply assign and assign back the variables, and export them into
the main namespace.

Note: In order to use this feature you may have to make changes to the registry.
This is at your own risk. If your computer stops working after you make these
changes I am not responsible. I have made the changes myself and they appear to
work, but please be careful.

Change the following keys:

	HKEY_LOCAL_MACHINE\
	  SYSTEM\
	    CurrentControlSet\
	      Services\
	        W3SVC\
	          ASP\
	            LanguageEngines\
	              PerlScript

	Write = "$main::Response->write(|);"
	WriteBlock = "$main::Response->writeblock(|);"

The change should be just to add the "main::" after the "$" and before "Response".

=item Print LIST

Prints a string or comma separated list of strings to the browser. Use
as if you were using print in a CGI application. Print gets around ASP's
limitations of 128k in a single Response->Write call.

Obsolete - use C<print> instead.

NB: C<print> calls Print, so you could use either, but print is more integrated
with "the perl way".

=cut
sub Print (@) {
	my ($output);
	foreach $output (@_) {
		if (length($output) > 128000) {
			Print (unpack('a128000a*', $output));
		}
		else {$main::Response->Write($output);}
	}
}

=item DebugPrint LIST

The same as C<Print> except the output is between HTML comments
so that you can only see it with "view source". DebugPrint is
not exported so you have to use it as Win32::ASP::DebugPrint()

This function is useful to debug your application. For example I
use it to print out SQL before it is executed.

=cut
sub DebugPrint (@) {
	Print "<!-- ";
	Print @_;
	Print " -->\n";
}

=item HTMLPrint LIST

The same as C<Print> except the output is taken and encoded so that
any html tags appear as sent, i.e. < becomes &lt;, > becomes &gt; etc.
HTMLPrint is not exported, so use it like Win32::ASP::HTMLPrint.

This function is useful for printing output that comes from a database
or a file, where you don't have total control over the input.

=cut
sub HTMLPrint (@) {
	my ($output);
	foreach $output (@_) {
		Print $Server->HTMLEncode($output);
	}
}

=item wprint LIST

Obsolete: Use C<Print> instead

=cut
sub wprint (@) {
	Print @_;
}

=item die LIST

Prints the contents of LIST to the browser and then exits. C<die> automatically
calls $Response->End for you, it also executes any cleanup code you have
added with C<AddDeathHook>.

=cut
sub die (@) {
	Print @_;
	Print "</BODY></HTML>";
	END;
	$main::Response->End();
	CORE::die();
}

=item exit

Exits the current script. $Response->End is called automatically for you, and
any cleanup code added with C<AddDeathHook> is also called.

=cut
sub exit (;$) {
	END;
	$main::Response->End();
	CORE::exit();
}

=item GetFormValue EXPR [, EXPR]

returns the value passed from a form (or non-form GET request). Use this
method if you want to be able to develop in GET mode (for ease of debugging)
and move to POST mode for release. The second (optional) parameter is for
getting multiple parameters as in:

	http://localhost/scripts/test.asp?Q=a&Q=b

In the above GetFormValue("Q", 1) returns "a" and GetFormValue("Q", 2)
returns "b".

GetFormValue will work in an array context too, returning all the values
for a particular parameter. For example with the above url:

	my @AllQs = GetFormValue('Q');

will return an array: @AllQs = ['a', 'b']

=cut
sub GetFormValue ($;$) {
	$_[1] = 1 unless defined $_[1];
	if (!wantarray) {
		if ($main::Request->ServerVariables('REQUEST_METHOD')->Item eq 'GET') {
			return $main::Request->QueryString($_[0])->Item($_[1]);
		}
		else {
			return $main::Request->Form($_[0])->Item($_[1]);
		}
	}
	else {
		my ($i, @ret);
		if ($main::Request->ServerVariables('REQUEST_METHOD')->Item eq 'GET') {
			my $count = $main::Request->QueryString($_[0])->{Count};
			for ($i = 1; $i <= $count; $i++ ) {
				push @ret, $main::Request->QueryString($_[0])->Item($i);
			}
		}
		else {
			my $count = $main::Request->Form($_[0])->{Count};
			for ($i = 1; $i <= $count; $i++) {
				push @ret, $main::Request->Form($_[0])->Item($i);
			}
		}
		return @ret;
	}
}

=item GetFormCount EXPR

returns the number of times EXPR appears in the request (Form or QueryString).
Use this value as $i to iterate over GetFormValue(EXPR, $i).

For example, url is:

	http://localhost/scripts/myscript.asp?Q=a&Q=b

And code is:

	my $numQs = GetFormCount('Q');

gives $numQs = 2

=cut
sub GetFormCount ($) {
	if ($main::Request->ServerVariables('REQUEST_METHOD')->Item eq 'GET') {
		return $main::Request->QueryString($_[0])->Count;
	}
	else {
		return $main::Request->Form($_[0])->Count;
	}
}

=item AddDeathHook LIST

This frightening sounding function allows you to have cleanup code
executed when you C<die> or C<exit>. For example you may want to
disconnect from your database if there is a problem:

	<%
		my $Conn = $Server->CreateObject('ADODB.Connection');
		$Conn->Open( "DSN=BADEV1;UID=sa;DATABASE=ProjAlloc" );
		$Conn->BeginTrans();

		Win32::ASP::AddDeathHook( sub { $Conn->Close if $Conn; } );
	%>

Now when you C<die> because of an error, your database connection
will close gracefully, instead of you having loads of rogue connections
that you have to kill by hand, or restart your database once a day.

Death hooks should be executed on a graceful exit of the script too,
but I've been unable to confirm this. If anyone has any luck, let me know.

=cut

sub AddDeathHook (@) {
	push @DeathHooks, @_;
}

END {
	my $func;
	foreach $func (@DeathHooks) {
		&$func();
	}
}

=item BinaryWrite LIST

Performs the same function as C<$Response->E<gt>C<BinaryWrite()> but gets around
Perl's lack of unicode support, and the null padding it uses to get around
this.

Example:

	Win32::ASP::BinaryWrite($val);

=cut

use Win32::OLE::Variant;

sub BinaryWrite (@) {
	my ($output);
	foreach $output (@_) {
		if (length($output) > 128000) {
			BinaryWrite (unpack('a128000a*', $output));
		}
		else {
			my $variant = Win32::OLE::Variant->new( VT_UI1, $output );
			$main::Response->BinaryWrite($variant);
		}
	}
}

# These two functions are ripped from CGI.pm
sub expire_calc {
    my($time) = @_;
    my(%mult) = ('s'=>1,
                 'm'=>60,
                 'h'=>60*60,
                 'd'=>60*60*24,
                 'M'=>60*60*24*30,
                 'y'=>60*60*24*365);
    # format for time can be in any of the forms...
    # "now" -- expire immediately
    # "+180s" -- in 180 seconds
    # "+2m" -- in 2 minutes
    # "+12h" -- in 12 hours
    # "+1d"  -- in 1 day
    # "+3M"  -- in 3 months
    # "+2y"  -- in 2 years
    # "-3m"  -- 3 minutes ago(!)
    # If you don't supply one of these forms, we assume you are
    # specifying the date yourself
    my($offset);
    if (!$time || ($time eq 'now')) {
        $offset = 0;
    } elsif ($time=~/^([+-]?\d+)([mhdMy]?)/) {
        $offset = ($mult{$2} || 1)*$1;
    } else {
        return $time;
    }
    return (time+$offset);
}

sub date {
    my($time,$format) = @_;
    my(@MON)=qw/Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec/;
    my(@WDAY) = qw/Sun Mon Tue Wed Thu Fri Sat/;

    # pass through preformatted dates for the sake of expire_calc()
    if ("$time" =~ m/^[^0-9]/o) {
        return $time;
    }

    # make HTTP/cookie date string from GMT'ed time
    # (cookies use '-' as date separator, HTTP uses ' ')
    my($sc) = ' ';
    $sc = '-' if $format eq "cookie";
    my($sec,$min,$hour,$mday,$mon,$year,$wday) = gmtime($time);
    $year += 1900;
    return sprintf("%s, %02d$sc%s$sc%04d %02d:%02d:%02d GMT",
                   $WDAY[$wday],$mday,$MON[$mon],$year,$hour,$min,$sec);
}

=item SetCookie Name, Value [, HASH]

Sets the cookie Name with the value Value. HASH is option, and contains any of
the following optional parameters:

=over 4

=item * -expires => A CGI.pm style expires value (see the CGI.pm docs for header() for this).

=item * -domain => a domain in the style ".matt.com" that the cookie is returned to.

=item * -path => a path that the cookie is returned to.

=item * -secure => cookie only gets returned under SSL if this is true.

=back

Example:

	Win32::ASP::SetCookie("Matt", "Sergeant", ( -expires => "+3h",
		-domain => ".matt.com",
		-path => "/users/matt",
		-secure => 0 ));

=cut

sub SetCookie ($$;%) {
	my ($name, $value, %hash) = @_;
	$main::Response->AddHeader('Set-Cookie', "$name=$value" .
		($hash{-path} ? "; path=" . $hash{-path} : "") .
		($hash{-domain} ? "; domain=" . $hash{-domain} : "") .
		($hash{-secure} ? "; secure" : "") .
		($hash{-expires} ? "; expires=" . &date(&expire_calc($hash{-expires})) : "") );
}

# print overloading

package Win32::ASP::IO;

sub new {
    my $self = bless {}, shift;
    $self;
}

sub print {
    my $self = shift;
    Win32::ASP::Print(@_);
    1;
}

sub TIEHANDLE { shift->new(@_) }
sub PRINT     { shift->print(@_) }
sub PRINTF    { shift->print(sprintf(@_)) }

1;
