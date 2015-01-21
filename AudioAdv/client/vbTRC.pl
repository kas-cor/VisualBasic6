#!/usr/bin/perl
##########################################################################
#                                            Winamp TCP/IP Remote Control																		   #
#                                                                   vbTRC                                                                                           #
#                                                                  v. 1.2.64                                                                                         #
#                                             Max V. Irgiznov vsBB   (c) 2002																		    #
#                                                       xeonvs@hotmail.com                                                                                 #
##########################################################################

use IO::Socket; 

$| = 1;  # Ќемедленный вывод следующей строки
#					0	 1		  2	
#usage: winamp.pl host <port> <command>

my @param=@ARGV; #parametr`s of command string

if (@param[0] ne ""){
	if (@param[2] ne "") {
		if (@param[1] ne "") {
			$cmd=@param[2];
			$port=@param[1];
		}#1
	} #2
	else {
		if (@param[1] ne "") {
			if ((@param[1] ne "info") && (@param[1] ne "play") && (@param[1] ne "stop") && (@param[1] ne "previous") && (@param[1] ne "next") && (@param[1] ne "pause") && (@param[1] ne "shuffle") && (@param[1] ne "np") && (substr(@param[1],0,1) ne "+") && (substr(@param[1],0,1) ne "-") ){
				if (@param[1] gt "A") {
					print "Error command! \n\n";
					print "usage: winamp.pl host <port> <command>; \n\n       Commands: info(default if <command> is NULL),\n        play, stop, previous, next, pause, shuffle, +(-)0..100 set volume\n"; 
					print "       Port default is 805\n\n";
					print "vsBBS (c) Xeon 2003\n";
					exit 0;
				} #port
				else {
					$port=@param[1];
					$cmd="np";
				} #e port
			}#cmd
			else {
				$cmd=@param[1];
				$port=805;
			} #e cmd
		} # 1
		else {
			$cmd="np";	
			$port=805;
		} # e1
	} #e2	
} # all
else {
	print "usage: winamp.pl host <port> <command>; \n\n       Commands: info(default if <command> is NULL),\n        play, stop, previous, next, pause, shuffle, +(-)0..100 set volume\n"; 
	print "       Port default is 805\n\n";
	print "vsBBS (c) Xeon 2002\n";
	exit 0;
} # all

$remote_host=@param[0];
$remote_port=$port;

#opening socket for work
print "Connect to $remote_host:$remote_port\n";
$socket = IO::Socket::INET->new( PeerAddr  => $remote_host,
				 PeerPort  => $remote_port,
			 	 Proto     => "tcp",
				 Type      => SOCK_STREAM)
or die "Couldn't connect to $remote_host:$remote_port: $@\n"; 

#Send Command
print "Send command: $cmd\n";
$socket->send($cmd, $flags)
	or die "Can't send: $!\n"; 

#Get string from socket
print "Now playing: ".<$socket>."\n";

#and exit

