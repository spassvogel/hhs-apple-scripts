FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � �
	will take an excel file and export all sheets as seperate pdfs
	using the first 8 characters of 
	
	
	
	NOTE: actually not working, todo!
     � 	 	 
 	 w i l l   t a k e   a n   e x c e l   f i l e   a n d   e x p o r t   a l l   s h e e t s   a s   s e p e r a t e   p d f s 
 	 u s i n g   t h e   f i r s t   8   c h a r a c t e r s   o f   
 	 
 	 
 	 
 	 N O T E :   a c t u a l l y   n o t   w o r k i n g ,   t o d o ! 
   
  
 l    	 ����  r     	    l     ����  I    ���� 
�� .sysostdfalis    ��� null��    �� ��
�� 
prmp  m       �   * S e l e c t   t h e   e x c e l   f i l e��  ��  ��    o      ���� 0 pathinputfile pathInputFile��  ��        l     ��  ��    W Qset destinationFolder to (choose folder with prompt "Where are the assessments?")     �   � s e t   d e s t i n a t i o n F o l d e r   t o   ( c h o o s e   f o l d e r   w i t h   p r o m p t   " W h e r e   a r e   t h e   a s s e s s m e n t s ? " )      l  
  ����  r   
     l  
  ����  I  
 ��   
�� .earsffdralis        afdr  m   
 ��
�� afdrdesk   �� !��
�� 
rtyp ! m    ��
�� 
ctxt��  ��  ��    o      ���� &0 destinationfolder destinationFolder��  ��     " # " l     ��������  ��  ��   #  $�� $ l   � %���� % O    � & ' & k    � ( (  ) * ) I   ������
�� .miscactvnull��� ��� null��  ��   *  + , + I   &�� -��
�� .aevtodocnull  �    alis - 4    "�� .
�� 
file . o     !���� 0 pathinputfile pathInputFile��   ,  / 0 / l  ' '�� 1 2��   1 M G set originalWorkbook to open workbook workbook file name pathInputFile    2 � 3 3 �   s e t   o r i g i n a l W o r k b o o k   t o   o p e n   w o r k b o o k   w o r k b o o k   f i l e   n a m e   p a t h I n p u t F i l e 0  4 5 4 r   ' , 6 7 6 l  ' * 8���� 8 b   ' * 9 : 9 o   ' (���� &0 destinationfolder destinationFolder : m   ( ) ; ; � < <  d a t a w o r l d . p d f��  ��   7 o      ���� 0 pdfname PDFName 5  = > = l  - -�� ? @��   ? G A save as active sheet filename myFile file format PDF file format    @ � A A �   s a v e   a s   a c t i v e   s h e e t   f i l e n a m e   m y F i l e   f i l e   f o r m a t   P D F   f i l e   f o r m a t >  B C B l  - -��������  ��  ��   C  D E D r   - 8 F G F n   - 4 H I H 2   0 4��
�� 
X128 I 1   - 0��
�� 
1172 G o      ���� 0 mysheets mySheets E  J K J l  9 9�� L M��   L  	log mySheets    M � N N  	 l o g   m y S h e e t s K  O�� O X   9 � P�� Q P O   O � R S R k   S � T T  U V U l  S S�� W X��   W C = set studentNumberAndName to get value of cell "B2" as string    X � Y Y z   s e t   s t u d e n t N u m b e r A n d N a m e   t o   g e t   v a l u e   o f   c e l l   " B 2 "   a s   s t r i n g V  Z [ Z r   S ` \ ] \ e   S \ ^ ^ c   S \ _ ` _ n   S X a b a 1   T X��
�� 
pnam b o   S T���� 0 asheet aSheet ` m   X [��
�� 
TEXT ] o      ���� ,0 studentnumberandname studentNumberAndName [  c d c l  a a��������  ��  ��   d  e�� e Z   a � f g���� f l  a m h���� h ?   a m i j i l  a i k���� k e   a i l l n   a i m n m 1   d h��
�� 
leng n o   a d���� ,0 studentnumberandname studentNumberAndName��  ��   j m   i l���� ��  ��   g k   p � o o  p q p l  p p�� r s��   r   log studentNumberAndName    s � t t 2   l o g   s t u d e n t N u m b e r A n d N a m e q  u v u r   p � w x w c   p � y z y n   p  { | { 7  s �� } ~
�� 
ctxt } m   w y����  ~ m   z ~����  | o   p s���� ,0 studentnumberandname studentNumberAndName z m    ���
�� 
TEXT x o      ���� 0 studentnumber studentNumber v   �  I  � ��� ���
�� .ascrcmnt****      � **** � o   � ����� 0 studentnumber studentNumber��   �  � � � l  � ���������  ��  ��   �  � � � O   � � � � � k   � � � �  � � � r   � � � � � I  � ����� �
�� .corecrel****      � null��   � �� ���
�� 
kocl � m   � ���
�� 
X141��   � o      ���� 0 newwbk newWBK �  � � � l  � ��� � ���   � 6 0activate object studentNumberAndName of mySheets    � � � � ` a c t i v a t e   o b j e c t   s t u d e n t N u m b e r A n d N a m e   o f   m y S h e e t s �  � � � l  � ��� � ���   � 5 / copy worksheet aSheet before sheet 1 of newWBK    � � � � ^   c o p y   w o r k s h e e t   a S h e e t   b e f o r e   s h e e t   1   o f   n e w W B K �  � � � l  � � � � � � 4   � ��� �
�� 
alis � o   � ����� 0 pdfname PDFName � * $ Necessary to avoid Excel Sandboxing    � � � � H   N e c e s s a r y   t o   a v o i d   E x c e l   S a n d b o x i n g �  � � � I  � ��� ���
�� .ascrcmnt****      � **** � o   � ����� 0 asheet aSheet��   �  ��� � l  � ��� � ���   � / )save aSheet in PDFName as PDF file format    � � � � R s a v e   a S h e e t   i n   P D F N a m e   a s   P D F   f i l e   f o r m a t��   � m   � � � ��                                                                                  XCEL  alis    F  Macintosh HD                   BD ����Microsoft Excel.app                                            ����            ����  
 cu             Applications  #/:Applications:Microsoft Excel.app/   (  M i c r o s o f t   E x c e l . a p p    M a c i n t o s h   H D   Applications/Microsoft Excel.app  / ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � 6 0			set AppleScript's text item delimiters to " "    � � � � ` 	 	 	 s e t   A p p l e S c r i p t ' s   t e x t   i t e m   d e l i m i t e r s   t o   "   " �  � � � l  � ��� � ���   � + %			log item 1 of studentNumberAndName    � � � � J 	 	 	 l o g   i t e m   1   o f   s t u d e n t N u m b e r A n d N a m e �  � � � l  � ��� � ���   � [ U log text 1 thru ((offset of "" in studentNumberAndName) - 1) of studentNumberAndName    � � � � �   l o g   t e x t   1   t h r u   ( ( o f f s e t   o f   " "   i n   s t u d e n t N u m b e r A n d N a m e )   -   1 )   o f   s t u d e n t N u m b e r A n d N a m e �  � � � l  � ��� � ���   � + %log "name is " & studentNumberAndName    � � � � J l o g   " n a m e   i s   "   &   s t u d e n t N u m b e r A n d N a m e �  ��� � l  � ���������  ��  ��  ��  ��  ��  ��   S o   O P���� 0 asheet aSheet�� 0 asheet aSheet Q o   < ?���� 0 mysheets mySheets��   ' m     � ��                                                                                  XCEL  alis    F  Macintosh HD                   BD ����Microsoft Excel.app                                            ����            ����  
 cu             Applications  #/:Applications:Microsoft Excel.app/   (  M i c r o s o f t   E x c e l . a p p    M a c i n t o s h   H D   Applications/Microsoft Excel.app  / ��  ��  ��  ��       �� � ���   � ��
�� .aevtoappnull  �   � **** � �� ����� � ���
�� .aevtoappnull  �   � **** � k     � � �  
 � �   � �  $����  ��  ��   � ���� 0 asheet aSheet � !�� �������������� ���~�} ;�|�{�z�y�x�w�v�u�t�s�r�q�p�o�n�m�l�k�j
�� 
prmp
�� .sysostdfalis    ��� null�� 0 pathinputfile pathInputFile
�� afdrdesk
�� 
rtyp
�� 
ctxt
�� .earsffdralis        afdr�� &0 destinationfolder destinationFolder
� .miscactvnull��� ��� null
�~ 
file
�} .aevtodocnull  �    alis�| 0 pdfname PDFName
�{ 
1172
�z 
X128�y 0 mysheets mySheets
�x 
kocl
�w 
cobj
�v .corecnte****       ****
�u 
pnam
�t 
TEXT�s ,0 studentnumberandname studentNumberAndName
�r 
leng�q �p �o 0 studentnumber studentNumber
�n .ascrcmnt****      � ****
�m 
X141
�l .corecrel****      � null�k 0 newwbk newWBK
�j 
alis�� �*��l E�O���l E�O� �*j 
O*��/j O��%E�O*�,a -E` O �_ [a a l kh  � i�a ,a &E` O_ a ,Ea  K_ [�\[Zk\Za 2a &E` O_ j O� !*a a l E` O*a  �/EO�j OPUOPY hU[OY��U ascr  ��ޭ