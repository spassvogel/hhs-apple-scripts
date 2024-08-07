FasdUAS 1.101.10   ��   ��    k             l      ��  ��   ��
====================================================
   [CSV] How to Read CSV File into AppleScript List
====================================================

PURPOSE:
   �ۢ Provide a simple process to read a CSV file, and parse each line/row into a AS list
       for further use as needed by the programmer.
   �ۢ This is a tool/example for the AppleScript programmer, not the end-user of a script
   �ۢ See limitations below
   
DATE:    Fri, Dec 11, 2015            VER: 1.0
AUTHOR: JMichaelTX (on MacScripter.net and StackOverflow.com forums)

LIMITATIONS:
   �ۢ I don't believe the parsing of the CSV data here conforms to the RFC 4180 de facto standard
   �ۢ It doesn't allow for clear separation of strings from numbers
   �ۢ If you need a more full-featured, compliant CSV parser, then see the below ref.

REF:    For a more robust, extensive process, see:
           CSV-to-list converter by Nigel Garvey, 2010-03-13
           [url]http://macscripter.net/viewtopic.php?pid=125444#p125444[/url]
           
SAMPLE INPUT DATA (from CSV file):
   Parent,Tag,Num Notes            <=== first line/row is column titles
   !SYMBOLS,SYM.ES,TBD
   ZIP_List,ZIP.77077,TBD
   FINANCE,FIN.Call,TBD
   Evernote,EN.UI,TBD
   HISTORY,HiS.NatlSec,TBD
   . . .

SAMPLE OUTPUT DATA (log)
   (*Number of Rows: 102*)
   (*!SYMBOLS, SYM.ES, TBD*)
   (*ZIP_List, ZIP.77077, TBD*)
   (*FINANCE, FIN.Call, TBD*)
   (*Evernote, EN.UI, TBD*)
   (*HISTORY, HiS.NatlSec, TBD*)
   . . .
====================================================
     � 	 	� 
 = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
       [ C S V ]   H o w   t o   R e a d   C S V   F i l e   i n t o   A p p l e S c r i p t   L i s t 
 = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 P U R P O S E : 
       � � �   P r o v i d e   a   s i m p l e   p r o c e s s   t o   r e a d   a   C S V   f i l e ,   a n d   p a r s e   e a c h   l i n e / r o w   i n t o   a   A S   l i s t 
               f o r   f u r t h e r   u s e   a s   n e e d e d   b y   t h e   p r o g r a m m e r . 
       � � �   T h i s   i s   a   t o o l / e x a m p l e   f o r   t h e   A p p l e S c r i p t   p r o g r a m m e r ,   n o t   t h e   e n d - u s e r   o f   a   s c r i p t 
       � � �   S e e   l i m i t a t i o n s   b e l o w 
       
 D A T E :         F r i ,   D e c   1 1 ,   2 0 1 5                         V E R :   1 . 0 
 A U T H O R :   J M i c h a e l T X   ( o n   M a c S c r i p t e r . n e t   a n d   S t a c k O v e r f l o w . c o m   f o r u m s ) 
 
 L I M I T A T I O N S : 
       � � �   I   d o n ' t   b e l i e v e   t h e   p a r s i n g   o f   t h e   C S V   d a t a   h e r e   c o n f o r m s   t o   t h e   R F C   4 1 8 0   d e   f a c t o   s t a n d a r d 
       � � �   I t   d o e s n ' t   a l l o w   f o r   c l e a r   s e p a r a t i o n   o f   s t r i n g s   f r o m   n u m b e r s 
       � � �   I f   y o u   n e e d   a   m o r e   f u l l - f e a t u r e d ,   c o m p l i a n t   C S V   p a r s e r ,   t h e n   s e e   t h e   b e l o w   r e f . 
 
 R E F :         F o r   a   m o r e   r o b u s t ,   e x t e n s i v e   p r o c e s s ,   s e e : 
                       C S V - t o - l i s t   c o n v e r t e r   b y   N i g e l   G a r v e y ,   2 0 1 0 - 0 3 - 1 3 
                       [ u r l ] h t t p : / / m a c s c r i p t e r . n e t / v i e w t o p i c . p h p ? p i d = 1 2 5 4 4 4 # p 1 2 5 4 4 4 [ / u r l ] 
                       
 S A M P L E   I N P U T   D A T A   ( f r o m   C S V   f i l e ) : 
       P a r e n t , T a g , N u m   N o t e s                         < = = =   f i r s t   l i n e / r o w   i s   c o l u m n   t i t l e s 
       ! S Y M B O L S , S Y M . E S , T B D 
       Z I P _ L i s t , Z I P . 7 7 0 7 7 , T B D 
       F I N A N C E , F I N . C a l l , T B D 
       E v e r n o t e , E N . U I , T B D 
       H I S T O R Y , H i S . N a t l S e c , T B D 
       .   .   . 
 
 S A M P L E   O U T P U T   D A T A   ( l o g ) 
       ( * N u m b e r   o f   R o w s :   1 0 2 * ) 
       ( * ! S Y M B O L S ,   S Y M . E S ,   T B D * ) 
       ( * Z I P _ L i s t ,   Z I P . 7 7 0 7 7 ,   T B D * ) 
       ( * F I N A N C E ,   F I N . C a l l ,   T B D * ) 
       ( * E v e r n o t e ,   E N . U I ,   T B D * ) 
       ( * H I S T O R Y ,   H i S . N a t l S e c ,   T B D * ) 
       .   .   . 
 = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
   
  
 l     ��  ��    - '- GET THE CSV FILE YOU WANT TO READ ---     �   N -   G E T   T H E   C S V   F I L E   Y O U   W A N T   T O   R E A D   - - -      l     ����  r         l    	 ����  I    	���� 
�� .sysostdfalis    ��� null��    ��  
�� 
prmp  m       �   & S e l e c t   t h e   C S V   f i l e  �� ��
�� 
ftyp  m       �    c s v��  ��  ��    o      ���� 0 pathinputfile pathInputFile��  ��        l     ��   ��    � � set _mailSubjectResponse to display dialog "Mail subject" default answer "Resultaten voor <VAK>" with icon note buttons {"Cancel", "Continue"} default button "Continue"      � ! !R   s e t   _ m a i l S u b j e c t R e s p o n s e   t o   d i s p l a y   d i a l o g   " M a i l   s u b j e c t "   d e f a u l t   a n s w e r   " R e s u l t a t e n   v o o r   < V A K > "   w i t h   i c o n   n o t e   b u t t o n s   { " C a n c e l " ,   " C o n t i n u e " }   d e f a u l t   b u t t o n   " C o n t i n u e "   " # " l     ��������  ��  ��   #  $ % $ l     �� & '��   &� set _mailBodyResponse to display dialog "Mail body" default answer "Beste {1}, " & linefeed & "bij dezen jouw resultaat voor proeftoets 2: {3}. " & linefeed & "Voor beide proeftoetsen heb je gemiddeld een: {4}. " & linefeed & linefeed & "met vriendelijke groet," & linefeed & linefeed & "Wouter van den Heuvel" with icon note buttons {"Cancel", "Continue"} default button "Continue"    ' � ( (�   s e t   _ m a i l B o d y R e s p o n s e   t o   d i s p l a y   d i a l o g   " M a i l   b o d y "   d e f a u l t   a n s w e r   " B e s t e   { 1 } ,   "   &   l i n e f e e d   &   " b i j   d e z e n   j o u w   r e s u l t a a t   v o o r   p r o e f t o e t s   2 :   { 3 } .   "   &   l i n e f e e d   &   " V o o r   b e i d e   p r o e f t o e t s e n   h e b   j e   g e m i d d e l d   e e n :   { 4 } .   "   &   l i n e f e e d   &   l i n e f e e d   &   " m e t   v r i e n d e l i j k e   g r o e t , "   &   l i n e f e e d   &   l i n e f e e d   &   " W o u t e r   v a n   d e n   H e u v e l "   w i t h   i c o n   n o t e   b u t t o n s   { " C a n c e l " ,   " C o n t i n u e " }   d e f a u l t   b u t t o n   " C o n t i n u e " %  ) * ) l     �� + ,��   + f ` set _mailbodyTemplate to replace_chars((text returned of _mailBodyResponse), linefeed, "<br/>")    , � - - �   s e t   _ m a i l b o d y T e m p l a t e   t o   r e p l a c e _ c h a r s ( ( t e x t   r e t u r n e d   o f   _ m a i l B o d y R e s p o n s e ) ,   l i n e f e e d ,   " < b r / > " ) *  . / . l     ��������  ��  ��   /  0 1 0 l     �� 2 3��   2 " - READ THE FILE CONTENTS ---    3 � 4 4 8 -   R E A D   T H E   F I L E   C O N T E N T S   - - - 1  5 6 5 l    7���� 7 r     8 9 8 I   �� :��
�� .rdwrread****        **** : o    ���� 0 pathinputfile pathInputFile��   9 o      ���� "0 strfilecontents strFileContents��  ��   6  ; < ; l     ��������  ��  ��   <  = > = l     �� ? @��   ? @ :- BREAK THE FILE INTO PARAGRAPHS (i.e., ROWS or LINES) ---    @ � A A t -   B R E A K   T H E   F I L E   I N T O   P A R A G R A P H S   ( i . e . ,   R O W S   o r   L I N E S )   - - - >  B C B l     �� D E��   D 7 1        (AS Paragraphs are separated by LF or CR)    E � F F b                 ( A S   P a r a g r a p h s   a r e   s e p a r a t e d   b y   L F   o r   C R ) C  G H G l     ��������  ��  ��   H  I J I l    K���� K r     L M L l    N���� N n     O P O 2   ��
�� 
cpar P o    ���� "0 strfilecontents strFileContents��  ��   M o      ���� "0 parfilecontents parFileContents��  ��   J  Q R Q l   ! S���� S r    ! T U T I   �� V��
�� .corecnte****       **** V o    ���� "0 parfilecontents parFileContents��   U o      ���� 0 numrows numRows��  ��   R  W X W l  " ) Y���� Y I  " )�� Z��
�� .ascrcmnt****      � **** Z b   " % [ \ [ m   " # ] ] � ^ ^   N u m b e r   o f   R o w s :   \ o   # $���� 0 numrows numRows��  ��  ��   X  _ ` _ l  * ; a���� a r   * ; b c b I   * 7�� d���� 0 parsecsv parseCSV d  e�� e c   + 3 f g f n   + / h i h 4   , /�� j
�� 
cobj j m   - .����  i o   + ,���� "0 parfilecontents parFileContents g m   / 2��
�� 
ctxt��  ��   c o      ���� &0 lstfieldsinheader lstFieldsinHeader��  ��   `  k l k l  < C m���� m I  < C�� n��
�� .ascrcmnt****      � **** n o   < ?���� &0 lstfieldsinheader lstFieldsinHeader��  ��  ��   l  o p o l     ��������  ��  ��   p  q r q l     �� s t��   s B <- PROCESS EACH PARAGRAPH (AKA LINE or ROW) OF INPUT FILE ---    t � u u x -   P R O C E S S   E A C H   P A R A G R A P H   ( A K A   L I N E   o r   R O W )   O F   I N P U T   F I L E   - - - r  v w v l  D a x���� x I  D a�� y z
�� .gtqpchltns    @   @ ns   y o   D G���� &0 lstfieldsinheader lstFieldsinHeader z �� { |
�� 
appr { m   J M } } � ~ ~  E m a i l   c o l u m n | ��  �
�� 
prmp  m   N Q � � � � � 4 P l e a s e   c h o o s e   e m a i l   c o l u m n � �� � �
�� 
inSL � m   T W � � � � �  E - m a i l a d r e s � �� ���
�� 
mlsl � m   Z [��
�� boovtrue��  ��  ��   w  � � � l      �� � ���   ���-- repeat with iPar from 2 to 3repeat with iPar from 2 to number of items in parFileContents	--        Skip first row since it has column titles, data starts in 2nd row		set lstRow to item iPar of parFileContents	if lstRow = "" then exit repeat -- EXIT Loop if Row is empty, like the last line		set lstFieldsinRow to parseCSV(lstRow as text)	set _mailbody to _mailbodyTemplate	set _recipient to item 2 of lstFieldsinRow		(*
	todo: allow index to be input by user, 
	make suggestions based on column names
	*)	set strName to item 5 of lstFieldsinRow -- COL 1 of CSV file	set strGrade to item 3 of lstFieldsinRow -- COL 2 of CSV file	set strGrade2 to item 4 of lstFieldsinRow -- COL 2 of CSV file	set numNotes to item 3 of lstFieldsinRow -- COL 3 of CSV file		repeat with i from 1 to count of lstFieldsinRow		set _mailbody to replace_chars(_mailbody, "{" & i & "}", item i of lstFieldsinRow)	end repeat	log _recipient		tell application "Microsoft Outlook"		set _message to make new outgoing message with properties {subject:(text returned of _mailSubjectResponse), content:_mailbody}		tell _message			make new recipient at _message with properties {email address:{address:_recipient}}		end tell		open _message	end tell	end repeat -- with iPar    � � � �	�  - -   r e p e a t   w i t h   i P a r   f r o m   2   t o   3  r e p e a t   w i t h   i P a r   f r o m   2   t o   n u m b e r   o f   i t e m s   i n   p a r F i l e C o n t e n t s  	 - -                 S k i p   f i r s t   r o w   s i n c e   i t   h a s   c o l u m n   t i t l e s ,   d a t a   s t a r t s   i n   2 n d   r o w  	  	 s e t   l s t R o w   t o   i t e m   i P a r   o f   p a r F i l e C o n t e n t s  	 i f   l s t R o w   =   " "   t h e n   e x i t   r e p e a t   - -   E X I T   L o o p   i f   R o w   i s   e m p t y ,   l i k e   t h e   l a s t   l i n e  	  	 s e t   l s t F i e l d s i n R o w   t o   p a r s e C S V ( l s t R o w   a s   t e x t )  	 s e t   _ m a i l b o d y   t o   _ m a i l b o d y T e m p l a t e  	 s e t   _ r e c i p i e n t   t o   i t e m   2   o f   l s t F i e l d s i n R o w  	  	 ( * 
 	 t o d o :   a l l o w   i n d e x   t o   b e   i n p u t   b y   u s e r ,   
 	 m a k e   s u g g e s t i o n s   b a s e d   o n   c o l u m n   n a m e s 
 	 * )  	 s e t   s t r N a m e   t o   i t e m   5   o f   l s t F i e l d s i n R o w   - -   C O L   1   o f   C S V   f i l e  	 s e t   s t r G r a d e   t o   i t e m   3   o f   l s t F i e l d s i n R o w   - -   C O L   2   o f   C S V   f i l e  	 s e t   s t r G r a d e 2   t o   i t e m   4   o f   l s t F i e l d s i n R o w   - -   C O L   2   o f   C S V   f i l e  	 s e t   n u m N o t e s   t o   i t e m   3   o f   l s t F i e l d s i n R o w   - -   C O L   3   o f   C S V   f i l e  	  	 r e p e a t   w i t h   i   f r o m   1   t o   c o u n t   o f   l s t F i e l d s i n R o w  	 	 s e t   _ m a i l b o d y   t o   r e p l a c e _ c h a r s ( _ m a i l b o d y ,   " { "   &   i   &   " } " ,   i t e m   i   o f   l s t F i e l d s i n R o w )  	 e n d   r e p e a t  	 l o g   _ r e c i p i e n t  	  	 t e l l   a p p l i c a t i o n   " M i c r o s o f t   O u t l o o k "  	 	 s e t   _ m e s s a g e   t o   m a k e   n e w   o u t g o i n g   m e s s a g e   w i t h   p r o p e r t i e s   { s u b j e c t : ( t e x t   r e t u r n e d   o f   _ m a i l S u b j e c t R e s p o n s e ) ,   c o n t e n t : _ m a i l b o d y }  	 	 t e l l   _ m e s s a g e  	 	 	 m a k e   n e w   r e c i p i e n t   a t   _ m e s s a g e   w i t h   p r o p e r t i e s   { e m a i l   a d d r e s s : { a d d r e s s : _ r e c i p i e n t } }  	 	 e n d   t e l l  	 	 o p e n   _ m e s s a g e  	 e n d   t e l l  	  e n d   r e p e a t   - -   w i t h   i P a r  �  � � � l     �� � ���   � 7 1=============== END OF MAIN SCRIPT ==============    � � � � b = = = = = = = = = = = = = = =   E N D   O F   M A I N   S C R I P T   = = = = = = = = = = = = = = �  � � � l     ��������  ��  ��   �  � � � i      � � � I      �� ����� 0 parsecsv parseCSV �  ��� � o      ���� 0 pstrrowtext pstrRowText��  ��   � k     & � �  � � � r      � � � J      � �  � � � n     � � � 1    ��
�� 
txdl �  f      �  ��� � m     � � � � �  ;��   � J       � �  � � � o      ���� 0 od   �  ��� � n      � � � 1    ��
�� 
txdl �  f    ��   �  � � � r     � � � n     � � � 2   ��
�� 
citm � o    ���� 0 pstrrowtext pstrRowText � o      ���� 0 
parsedtext 
parsedText �  � � � r    # � � � o    ���� 0 od   � n      � � � 1     "��
�� 
txdl �  f      �  ��� � L   $ & � � o   $ %���� 0 
parsedtext 
parsedText��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � + %  set TestString to "1-2-3-5-8-13-21"    � � � � J     s e t   T e s t S t r i n g   t o   " 1 - 2 - 3 - 5 - 8 - 1 3 - 2 1 " �  � � � l     �� � ���   � 2 , set myArray to my theSplit(TestString, "-")    � � � � X   s e t   m y A r r a y   t o   m y   t h e S p l i t ( T e s t S t r i n g ,   " - " ) �  � � � i     � � � I      �� ����� 0 thesplit theSplit �  � � � o      ���� 0 	thestring 	theString �  ��� � o      ���� 0 thedelimiter theDelimiter��  ��   � k      � �  � � � l     �� � ���   � . ( save delimiters to restore old settings    � � � � P   s a v e   d e l i m i t e r s   t o   r e s t o r e   o l d   s e t t i n g s �  � � � r      � � � n     � � � 1    ��
�� 
txdl � 1     ��
�� 
ascr � o      ���� 0 olddelimiters oldDelimiters �  � � � l   �� � ���   � - ' set delimiters to delimiter to be used    � � � � N   s e t   d e l i m i t e r s   t o   d e l i m i t e r   t o   b e   u s e d �  � � � r     � � � o    ���� 0 thedelimiter theDelimiter � n      � � � 1    
��
�� 
txdl � 1    ��
�� 
ascr �  � � � l   �� � ���   �   create the array    � � � � "   c r e a t e   t h e   a r r a y �  � � � r     � � � n     � � � 2    ��
�� 
citm � o    ���� 0 	thestring 	theString � o      ���� 0 thearray theArray �  � � � l   �� � ���   �   restore the old setting    � � � � 0   r e s t o r e   t h e   o l d   s e t t i n g �  � � � r     � � � o    �� 0 olddelimiters oldDelimiters � n      � � � 1    �~
�~ 
txdl � 1    �}
�} 
ascr �  � � � l   �| � ��|   �   return the result    � �   $   r e t u r n   t h e   r e s u l t � �{ L     o    �z�z 0 thearray theArray�{   �  l     �y�x�w�y  �x  �w    i     I      �v	�u�v 0 replace_chars  	 

 o      �t�t 	0 _text    o      �s�s 0 _search_string   �r o      �q�q 0 _replacement_string  �r  �u   k        r      l    �p�o o     �n�n 0 _search_string  �p  �o   n      1    �m
�m 
txdl 1    �l
�l 
ascr  r     n    	 2    	�k
�k 
citm o    �j�j 	0 _text   l     �i�h o      �g�g 0 	item_list  �i  �h    r     !  l   "�f�e" o    �d�d 0 _replacement_string  �f  �e  ! n     #$# 1    �c
�c 
txdl$ 1    �b
�b 
ascr %&% r    '(' c    )*) l   +�a�`+ o    �_�_ 0 	item_list  �a  �`  * m    �^
�^ 
TEXT( o      �]�] 0 
_this_text  & ,-, r    ./. m    00 �11  / n     232 1    �\
�\ 
txdl3 1    �[
�[ 
ascr- 4�Z4 L     55 o    �Y�Y 0 
_this_text  �Z   6�X6 l     �W�V�U�W  �V  �U  �X       �T789:;�T  7 �S�R�Q�P�S 0 parsecsv parseCSV�R 0 thesplit theSplit�Q 0 replace_chars  
�P .aevtoappnull  �   � ****8 �O ��N�M<=�L�O 0 parsecsv parseCSV�N �K>�K >  �J�J 0 pstrrowtext pstrRowText�M  < �I�H�G�I 0 pstrrowtext pstrRowText�H 0 od  �G 0 
parsedtext 
parsedText= �F ��E�D
�F 
txdl
�E 
cobj
�D 
citm�L ')�,�lvE[�k/E�Z[�l/)�,FZO��-E�O�)�,FO�9 �C ��B�A?@�@�C 0 thesplit theSplit�B �?A�? A  �>�=�> 0 	thestring 	theString�= 0 thedelimiter theDelimiter�A  ? �<�;�:�9�< 0 	thestring 	theString�; 0 thedelimiter theDelimiter�: 0 olddelimiters oldDelimiters�9 0 thearray theArray@ �8�7�6
�8 
ascr
�7 
txdl
�6 
citm�@ ��,E�O���,FO��-E�O���,FO�: �5�4�3BC�2�5 0 replace_chars  �4 �1D�1 D  �0�/�.�0 	0 _text  �/ 0 _search_string  �. 0 _replacement_string  �3  B �-�,�+�*�)�- 	0 _text  �, 0 _search_string  �+ 0 _replacement_string  �* 0 	item_list  �) 0 
_this_text  C �(�'�&�%0
�( 
ascr
�' 
txdl
�& 
citm
�% 
TEXT�2 !���,FO��-E�O���,FO��&E�O���,FO�; �$E�#�"FG�!
�$ .aevtoappnull  �   � ****E k     aHH  II  5JJ  IKK  QLL  WMM  _NN  kOO  v� �   �#  �"  F  G � � ��������� ]������ } �� ����
� 
prmp
� 
ftyp� 
� .sysostdfalis    ��� null� 0 pathinputfile pathInputFile
� .rdwrread****        ****� "0 strfilecontents strFileContents
� 
cpar� "0 parfilecontents parFileContents
� .corecnte****       ****� 0 numrows numRows
� .ascrcmnt****      � ****
� 
cobj
� 
ctxt� 0 parsecsv parseCSV� &0 lstfieldsinheader lstFieldsinHeader
� 
appr
� 
inSL
� 
mlsl� 
� .gtqpchltns    @   @ ns  �! b*����� E�O�j E�O��-E�O�j E�O��%j O*��k/a &k+ E` O_ j O_ a a �a a a a ea  ascr  ��ޭ