                      ___
                     |   |           ___
                     |   |          |   |
                     |   |__________|   |  _______
                     |        |  ___\   |__\   ___\
                     |___|    |  ___/   |_ \|  ___/
                         |    |_____>______/|   |
.-=----------------------|____|-------------|___|---------------------=-.
| ## CMD    LINE               FUNCTION                             FLAG|
|-=-------------------------------------------------------------------=-|
|                                                                       |
| USER      %U                 - Shows information about that user  *   |
| USERS                        - Show full list of all users        17  |
| ADDIP     %U %P              - Add an ip to the user              127 |
| DELIP     %U %Ix             - Del an ip from the user            127 |
| ADDUSER   %U %P %I1 %I2      - Add user with optional 2 ip's      127 |
| GADDUSER  %G %U %P % %I1 %I2 - Add user to grp with 2 opt. ip's   127 |
| DELUSER   %U                 - Mark a user as deleted '6'         127 |
| GIVE      %U %K              - Give KB to user                    G   |
| NUKE      %D %M %R           - Nuke a directory ?x with reason    A   |
| RENUSER   %U %U              - Rename an user                     127 |
| SEEN      %U                 - Check when a user was last online  *   |
| TAGLINE   %T                 - Either check or set your tagline   *   |
| VERS                         - Return the server version          *   |
| TIME                         - Return the server time             *   |
| READD     %U                 - Readd a user.                      127 |
| WHO                          - Return who is currently online     *   |
| KICK      %Uid               - Kick user (get ID from site who)   D   |
| PASSWD    %P                 - Change your password               *   |
| HELP                         - Display the HELP file              *   |
| WELCOME                      - Display the WELCOME file           *   |
| ONEL      %O[%T]             - Either show or add msg to board    *   |
| REQUEST   %Q                 - Will place a request in req.dir    *   |
| REQFILLED %Q                 - Will mark the req. as filled       *   |
| NEWS      %T                 - Shows the news -only flag 1+7 add  *   |
| MOUNTDIR  %T %D              - Mount a directory                  17  |
| MOUNTCLEAR %T                - Clear a mounted directory          17  |
|-=-------------------------------------------------------------------=-|
| %U = Username   %P = Password    %Ix = IPx         %G = Group,        |
| %R = Reason     %Q = Request     %O  = Optional    %M = Multiplier    |
| %K = KBytes     %D = Directory   %T  = Text                           |
|-=-------------------------------------------------------------------=-|
| FLAG     : 1 = SiteOP - 2 = Group Admin - 3 = CoSiteop - * = all      |
|          : A = Nuke   - G = Give        - D = Kick                    |
`-=-------------------------------------------------------------------=-'