FasdUAS 1.101.10   ��   ��    k             l     ��  ��      BibDesk WordScript     � 	 	 &   B i b D e s k   W o r d S c r i p t   
  
 l     ��������  ��  ��        l     ��  ��      MIT License     �      M I T   L i c e n s e      l     ��������  ��  ��        l     ��  ��    ' ! Copyright (c) 2020 David Wingate     �   B   C o p y r i g h t   ( c )   2 0 2 0   D a v i d   W i n g a t e      l     ��������  ��  ��        l     ��  ��    S M Permission is hereby granted, free of charge, to any person obtaining a copy     �   �   P e r m i s s i o n   i s   h e r e b y   g r a n t e d ,   f r e e   o f   c h a r g e ,   t o   a n y   p e r s o n   o b t a i n i n g   a   c o p y       l     �� ! "��   ! T N of this software and associated documentation files (the "Software"), to deal    " � # # �   o f   t h i s   s o f t w a r e   a n d   a s s o c i a t e d   d o c u m e n t a t i o n   f i l e s   ( t h e   " S o f t w a r e " ) ,   t o   d e a l    $ % $ l     �� & '��   & S M in the Software without restriction, including without limitation the rights    ' � ( ( �   i n   t h e   S o f t w a r e   w i t h o u t   r e s t r i c t i o n ,   i n c l u d i n g   w i t h o u t   l i m i t a t i o n   t h e   r i g h t s %  ) * ) l     �� + ,��   + P J to use, copy, modify, merge, publish, distribute, sublicense, and/or sell    , � - - �   t o   u s e ,   c o p y ,   m o d i f y ,   m e r g e ,   p u b l i s h ,   d i s t r i b u t e ,   s u b l i c e n s e ,   a n d / o r   s e l l *  . / . l     �� 0 1��   0 L F copies of the Software, and to permit persons to whom the Software is    1 � 2 2 �   c o p i e s   o f   t h e   S o f t w a r e ,   a n d   t o   p e r m i t   p e r s o n s   t o   w h o m   t h e   S o f t w a r e   i s /  3 4 3 l     �� 5 6��   5 ? 9 furnished to do so, subject to the following conditions:    6 � 7 7 r   f u r n i s h e d   t o   d o   s o ,   s u b j e c t   t o   t h e   f o l l o w i n g   c o n d i t i o n s : 4  8 9 8 l     ��������  ��  ��   9  : ; : l     �� < =��   < U O The above copyright notice and this permission notice shall be included in all    = � > > �   T h e   a b o v e   c o p y r i g h t   n o t i c e   a n d   t h i s   p e r m i s s i o n   n o t i c e   s h a l l   b e   i n c l u d e d   i n   a l l ;  ? @ ? l     �� A B��   A 6 0 copies or substantial portions of the Software.    B � C C `   c o p i e s   o r   s u b s t a n t i a l   p o r t i o n s   o f   t h e   S o f t w a r e . @  D E D l     ��������  ��  ��   E  F G F l     �� H I��   H Q K THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR    I � J J �   T H E   S O F T W A R E   I S   P R O V I D E D   " A S   I S " ,   W I T H O U T   W A R R A N T Y   O F   A N Y   K I N D ,   E X P R E S S   O R G  K L K l     �� M N��   M O I IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,    N � O O �   I M P L I E D ,   I N C L U D I N G   B U T   N O T   L I M I T E D   T O   T H E   W A R R A N T I E S   O F   M E R C H A N T A B I L I T Y , L  P Q P l     �� R S��   R R L FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE    S � T T �   F I T N E S S   F O R   A   P A R T I C U L A R   P U R P O S E   A N D   N O N I N F R I N G E M E N T .   I N   N O   E V E N T   S H A L L   T H E Q  U V U l     �� W X��   W M G AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER    X � Y Y �   A U T H O R S   O R   C O P Y R I G H T   H O L D E R S   B E   L I A B L E   F O R   A N Y   C L A I M ,   D A M A G E S   O R   O T H E R V  Z [ Z l     �� \ ]��   \ T N LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,    ] � ^ ^ �   L I A B I L I T Y ,   W H E T H E R   I N   A N   A C T I O N   O F   C O N T R A C T ,   T O R T   O R   O T H E R W I S E ,   A R I S I N G   F R O M , [  _ ` _ l     �� a b��   a T N OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE    b � c c �   O U T   O F   O R   I N   C O N N E C T I O N   W I T H   T H E   S O F T W A R E   O R   T H E   U S E   O R   O T H E R   D E A L I N G S   I N   T H E `  d e d l     �� f g��   f  
 SOFTWARE.    g � h h    S O F T W A R E . e  i j i l     ��������  ��  ��   j  k l k l     �� m n��   m , & check that Word and BibDesk are ready    n � o o L   c h e c k   t h a t   W o r d   a n d   B i b D e s k   a r e   r e a d y l  p q p l    N r���� r O     N s t s k    M u u  v w v r     x y x J     z z  { | { m    ����   |  }�� } m    ����  ��   y J       ~ ~   �  o      ���� 0 	countword 	countWord �  ��� � o      ���� 0 countbibdesk countBibDesk��   w  � � � Z    2 � ����� � =   " � � � l     ����� � I    �� ���
�� .coredoexnull���     **** � 4    �� �
�� 
prcs � m     � � � � �  M i c r o s o f t   W o r d��  ��  ��   � m     !��
�� boovtrue � r   % . � � � l  % , ����� � I  % ,�� ���
�� .corecnte****       **** � n  % ( � � � 2  & (��
�� 
docu � m   % & � ��                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ��  ��  ��   � o      ���� 0 	countword 	countWord��  ��   �  ��� � Z   3 M � ����� � =  3 = � � � l  3 ; ����� � I  3 ;�� ���
�� .coredoexnull���     **** � 4   3 7�� �
�� 
prcs � m   5 6 � � � � �  B i b D e s k��  ��  ��   � m   ; <��
�� boovtrue � r   @ I � � � l  @ G ����� � I  @ G�� ���
�� .corecnte****       **** � n  @ C � � � 2  A C��
�� 
docu � m   @ A � ��                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  ��  ��  ��   � o      ���� 0 countbibdesk countBibDesk��  ��  ��   t m      � ��                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ��  ��   q  � � � l  O � ����� � Z   O � � ����� � E   O U � � � J   O S � �  � � � o   O P���� 0 	countword 	countWord �  ��� � o   P Q���� 0 countbibdesk countBibDesk��   � m   S T����   � k   X � � �  � � � I  X ]������
�� .miscactvnull��� ��� null��  ��   �  ��� � Q   ^ � � � � � k   a � � �  � � � I  a u�� � �
�� .sysodlogaskr        TEXT � m   a b � � � � � r D o c u m e n t s   m u s t   b e   o p e n   i n   b o t h   M i c r o s o f t   W o r d   a n d   B i b D e s k � �� � �
�� 
btns � J   c i � �  � � � m   c d � � � � �  C a n c e l �  ��� � m   d g � � � � �  O p e n   N o w��   � �� ���
�� 
dflt � m   l o � � � � �  O p e n   N o w��   �  ��� � Z   v � � ����� � E   v � � � � n   v } � � � 1   y }��
�� 
bhit � 1   v y��
�� 
rslt � m   } � � � � � �  O p e n   N o w � k   � � � �  � � � O   � � � � � I  � �������
�� .ascrnoop****      � ****��  ��   � m   � � � ��                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��   �  ��� � I  � ��� ���
�� .sysodelanull��� ��� nmbr � m   � ����� ��  ��  ��  ��  ��   � R      ������
�� .ascrerr ****      � ****��  ��   � L   � �����  ��  ��  ��  ��  ��   �  � � � l     �� � ���   � : 4 check for selection in Word to trigger quick format    � � � � h   c h e c k   f o r   s e l e c t i o n   i n   W o r d   t o   t r i g g e r   q u i c k   f o r m a t �  � � � l  � � ����� � O   � � � � � k   � � � �  � � � Z   � � � ����� � E   � � � � � n   � � � � � 1   � ���
�� 
1650 � n   � � � � � 1   � ���
�� 
wTxR � n   � � � � � 4  � ��� �
�� 
cpar � m   � �����  � 1   � ���
�� 
sele � m   � � � � � � � 
 A D D I N � l  � � � � � � I  � ��� ���
�� .miscslctnull���    obj  � n   � � � � � 1   � ���
�� 
wTxR � n   � � � � � 4  � ��� �
�� 
cpar � m   � �����  � 1   � ���
�� 
sele��   � ' ! select recently edited reference    � �   B   s e l e c t   r e c e n t l y   e d i t e d   r e f e r e n c e��  ��   �  r   � � l  � ����� n   � � 1   � ���
�� 
2903 1   � ���
�� 
sele��  ��   o      ��  0 selectionstart selectionStart �~ r   � �	
	 l  � ��}�| n   � � 1   � ��{
�{ 
2905 1   � ��z
�z 
sele�}  �|  
 o      �y�y 0 selectionend selectionEnd�~   � m   � ��                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ��  ��   �  l  �[�x�w Z   �[�v�u >  � � o   � ��t�t  0 selectionstart selectionStart o   � ��s�s 0 selectionend selectionEnd k  W  I  �r�q�p�r 40 readbibliographysettings readBibliographySettings�q  �p    r   1  �o
�o 
rslt o      �n�n ,0 bibliographysettings bibliographySettings  Z  K �m�l = !"! o  �k�k ,0 bibliographysettings bibliographySettings" J  �j�j    k  G## $%$ I  �i�h�g�i "0 defaultsettings defaultSettings�h  �g  % &'& r  &()( 1  "�f
�f 
rslt) o      �e�e ,0 bibliographysettings bibliographySettings' *+* I  '/�d,�c�d 60 writebibliographysettings writeBibliographySettings, -�b- o  (+�a�a ,0 bibliographysettings bibliographySettings�b  �c  + ./. r  07010 1  03�`
�` 
rslt1 o      �_�_ ,0 bibliographysettings bibliographySettings/ 2�^2 Z  8G34�]�\3 = 8>565 o  8;�[�[ ,0 bibliographysettings bibliographySettings6 J  ;=�Z�Z  4 L  AC�Y�Y  �]  �\  �^  �m  �l   787 I  LT�X9�W�X $0 formatreferences formatReferences9 :�V: o  MP�U�U ,0 bibliographysettings bibliographySettings�V  �W  8 ;�T; L  UW�S�S  �T  �v  �u  �x  �w   <=< l     �R>?�R  >   open WordScript Menu   ? �@@ *   o p e n   W o r d S c r i p t   M e n u= ABA l \aC�Q�PC I \a�O�N�M
�O .miscactvnull��� ��� null�N  �M  �Q  �P  B DED l b�F�L�KF t  b�GHG k  d�II JKJ I d��JLM
�J .gtqpchltns    @   @ ns  L J  dwNN OPO m  dgQQ �RR " F o r m a t   R e f e r e n c e sP STS m  gjUU �VV & U n f o r m a t   R e f e r e n c e sT WXW m  jmYY �ZZ * F i l l   E m p t y   R e f e r e n c e sX [\[ m  mp]] �^^ 
 T o o l s\ _�I_ m  ps`` �aa " D o c u m e n t   S e t t i n g s�I  M �Hbc
�H 
apprb m  z}dd �ee  W o r d S c r i p tc �Gfg
�G 
prmpf m  ��hh �ii @ Y o u   s h o u l d   s a v e   y o u r   w o r k   f i r s t !g �Fjk
�F 
inSLj m  ��ll �mm " F o r m a t   R e f e r e n c e sk �Eno
�E 
cnbtn J  ��pp q�Dq m  ��rr �ss  E x i t�D  o �Ct�B
�C 
okbtt J  ��uu v�Av m  ��ww �xx  C h o o s e�A  �B  K y�@y r  ��z{z 1  ���?
�? 
rslt{ o      �>�> 0 
menuchoice 
menuChoice�@  H m  bc�=�=  �L  �K  E |}| l ��~�<�;~ Z  ����:�9 = ����� o  ���8�8 0 
menuchoice 
menuChoice� m  ���7
�7 boovfals� L  ���6�6  �:  �9  �<  �;  } ��� l �[��5�4� Z  �[����3� E  ����� o  ���2�2 0 
menuchoice 
menuChoice� m  ���� ��� " F o r m a t   R e f e r e n c e s� k  ��� ��� I  ���1�0�/�1 40 readbibliographysettings readBibliographySettings�0  �/  � ��� r  ����� 1  ���.
�. 
rslt� o      �-�- ,0 bibliographysettings bibliographySettings� ��� Z  ����,�+� = ����� o  ���*�* ,0 bibliographysettings bibliographySettings� J  ���)�)  � k  ��� ��� I  ���(�'�&�( "0 defaultsettings defaultSettings�'  �&  � ��� r  ����� 1  ���%
�% 
rslt� o      �$�$ ,0 bibliographysettings bibliographySettings� ��� I  ���#��"�# 60 writebibliographysettings writeBibliographySettings� ��!� o  ��� �  ,0 bibliographysettings bibliographySettings�!  �"  � ��� r  ����� 1  ���
� 
rslt� o      �� ,0 bibliographysettings bibliographySettings� ��� Z  ������ = ����� o  ���� ,0 bibliographysettings bibliographySettings� J  ����  � L  ��  �  �  �  �,  �+  � ��� I  ���� $0 formatreferences formatReferences� ��� o  �� ,0 bibliographysettings bibliographySettings�  �  �  � ��� E  ��� o  �� 0 
menuchoice 
menuChoice� m  �� ��� & U n f o r m a t   R e f e r e n c e s� ��� I  !&���� (0 unformatreferences unformatReferences�  �  � ��� E  )0��� o  ),�� 0 
menuchoice 
menuChoice� m  ,/�� ��� * F i l l   E m p t y   R e f e r e n c e s� ��� I  38���� *0 fillemptyreferences fillEmptyReferences�  �  � ��� E  ;B��� o  ;>�
�
 0 
menuchoice 
menuChoice� m  >A�� ��� " D o c u m e n t   S e t t i n g s� ��� k  E��� ��� I  EJ�	���	 40 readbibliographysettings readBibliographySettings�  �  � ��� r  KR��� 1  KN�
� 
rslt� o      �� ,0 bibliographysettings bibliographySettings� ��� Z  S������ = SY��� o  SV�� ,0 bibliographysettings bibliographySettings� J  VX��  � k  \��� ��� I  \a� �����  "0 defaultsettings defaultSettings��  ��  � ��� r  bi��� 1  be��
�� 
rslt� o      ���� ,0 bibliographysettings bibliographySettings� ��� I  jr������� 60 writebibliographysettings writeBibliographySettings� ���� o  kn���� ,0 bibliographysettings bibliographySettings��  ��  � ��� r  sz��� 1  sv��
�� 
rslt� o      ���� ,0 bibliographysettings bibliographySettings� ���� Z  {�������� = {���� o  {~���� ,0 bibliographysettings bibliographySettings� J  ~�����  � L  ������  ��  ��  ��  �  �  � ��� I  ��������� 0 settingsmenu settingsMenu� ���� o  ������ ,0 bibliographysettings bibliographySettings��  ��  � ���� Z  �������� > ����� 1  ����
�� 
rslt� o  ������ ,0 bibliographysettings bibliographySettings� k  ���� ��� r  ����� 1  ����
�� 
rslt� o      ���� ,0 bibliographysettings bibliographySettings� ��� I  ��������� 60 writebibliographysettings writeBibliographySettings� ���� o  ������ ,0 bibliographysettings bibliographySettings��  ��  � ��� r  ��� � 1  ����
�� 
rslt  o      ���� ,0 bibliographysettings bibliographySettings�  Z  ������ = �� o  ������ ,0 bibliographysettings bibliographySettings J  ������   L  ������  ��  ��   �� Q  ��	
 k  ��  O  �� k  ��  I ��������
�� .miscactvnull��� ��� null��  ��   �� I ������
�� .sysodlogaskr        TEXT m  �� � � S e t t i n g s   h a v e   b e e n   c h a n g e d ,   d o   y o u   w i s h   t o   f o r m a t   t h e   r e f e r e n c e s   n o w ?��  ��   m  ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��   �� I  �������� $0 formatreferences formatReferences �� o  ������ ,0 bibliographysettings bibliographySettings��  ��  ��  	 R      ������
�� .ascrerr ****      � ****��  ��  
 L  ������  ��  ��  � L  ������  ��  �  E  � o  ������ 0 
menuchoice 
menuChoice m  � �   
 T o o l s !��! k  W"" #$# Q  %&'% k  (( )*) I ��+��
�� .sysodlogaskr        TEXT+ m  ,, �-- x R e f e r e n c e s   m u s t   b e   u n f o r m a t t e d ,   d o   y o u   w i s h   t o   u n f o r m a t   n o w ?��  * .��. I  �������� (0 unformatreferences unformatReferences��  ��  ��  & R      ������
�� .ascrerr ****      � ****��  ��  ' L  ����  $ /0/ I   %�������� 0 	toolsmenu 	toolsMenu��  ��  0 1��1 Q  &W2342 k  )M55 676 I ).������
�� .miscactvnull��� ��� null��  ��  7 898 I /6��:��
�� .sysodlogaskr        TEXT: m  /2;; �<< z O p e r a t i o n   c o m p l e t e ,   d o   y o u   w i s h   t o   f o r m a t   t h e   r e f e r e n c e s   n o w ?��  9 =>= I  7<�������� 40 readbibliographysettings readBibliographySettings��  ��  > ?@? r  =DABA 1  =@��
�� 
rsltB o      ���� ,0 bibliographysettings bibliographySettings@ C��C I  EM��D���� $0 formatreferences formatReferencesD E��E o  FI���� ,0 bibliographysettings bibliographySettings��  ��  ��  3 R      ������
�� .ascrerr ****      � ****��  ��  4 L  UW����  ��  ��  �3  �5  �4  � FGF l     ��������  ��  ��  G HIH l     ��JK��  J  -   K �LL  -I MNM l     ��������  ��  ��  N OPO i     QRQ I      �������� "0 defaultsettings defaultSettings��  ��  R k     USS TUT r     VWV m     XX �YY  W o      ���� *0 bibdeskdocumentname bibDeskDocumentNameU Z[Z r    \]\ m    ^^ �__  ] o      ���� 40 bibliographytemplatepath bibliographyTemplatePath[ `a` r    bcb m    	dd �ee  c o      ���� &0 citeptemplatepath citepTemplatePatha fgf r    hih m    jj �kk  i o      ���� &0 citettemplatepath citetTemplatePathg lml r    non m    pp �qq  (o o      ���� "0 parenthesisopen parenthesisOpenm rsr r    tut m    vv �ww  )u o      ���� $0 parenthesisclose parenthesisCloses xyx l   z{|z r    }~} m     ���  C i t e   K e y~ o      ���� (0 bibliographysortby bibliographySortBy{ : 4 BibDesk field (leave blank for order of appearance)   | ��� h   B i b D e s k   f i e l d   ( l e a v e   b l a n k   f o r   o r d e r   o f   a p p e a r a n c e )y ��� l   ���� r    ��� m    �� ���  B i b l i o g r a p h y� o      ���� &0 bibliographystyle bibliographyStyle� 5 / leave blank to use formatting of template file   � ��� ^   l e a v e   b l a n k   t o   u s e   f o r m a t t i n g   o f   t e m p l a t e   f i l e� ��� r     )��� J     '�� ��� m     !�� ���  e t   a l .� ��� m   ! "�� ��� 
 i b i d .� ��� m   " #�� ���  c f .� ��� m   # $�� ���  i n t e r   a l i a� ���� m   $ %�� ��� 
 c i r c a��  � o      ���� 0 italiclatin italicLatin� ��� r   * 0��� J   * .�� ��� m   * +�� ���  f i l m� ���� m   + ,�� ���  b r o a d c a s t��  � o      ���� 00 italicpublicationtypes italicPublicationTypes� ��� r   1 6��� m   1 4�� ���  N o� o      ���� .0 superscriptreferences superscriptReferences� ��� l  7 <���� r   7 <��� m   7 :�� ���  N o� o      ���� (0 numberedreferences numberedReferences� &   overrides templates and sorting   � ��� @   o v e r r i d e s   t e m p l a t e s   a n d   s o r t i n g� ��� r   = B��� m   = @�� ���  ,� o      ���� "0 numberseperator numberSeperator� ���� L   C U�� J   C T�� ��� o   C D���� *0 bibdeskdocumentname bibDeskDocumentName� ��� o   D E���� 40 bibliographytemplatepath bibliographyTemplatePath� ��� o   E F���� &0 citeptemplatepath citepTemplatePath� ��� o   F G���� &0 citettemplatepath citetTemplatePath� ��� o   G H���� "0 parenthesisopen parenthesisOpen� ��� o   H I���� $0 parenthesisclose parenthesisClose� ��� o   I J���� (0 bibliographysortby bibliographySortBy� ��� o   J K���� &0 bibliographystyle bibliographyStyle� ��� o   K L���� 0 italiclatin italicLatin� ��� o   L M���� 00 italicpublicationtypes italicPublicationTypes� ��� o   M N���� .0 superscriptreferences superscriptReferences� ��� o   N O���� (0 numberedreferences numberedReferences� ���� o   O P�� "0 numberseperator numberSeperator��  ��  P ��� l     �~�}�|�~  �}  �|  � ��� l     �{���{  �  -   � ���  -� ��� l     �z�y�x�z  �y  �x  � ��� i    ��� I      �w�v�u�w 40 readbibliographysettings readBibliographySettings�v  �u  � k    >�� ��� l     �t���t  �   global setup   � ���    g l o b a l   s e t u p� ��� O     $��� k    #�� ��� r    	��� 1    �s
�s 
1003� o      �r�r 0 worddocument wordDocument� ��� r   
    I  
 �q
�q .sWRD1430null���     docu o   
 �p�p 0 worddocument wordDocument �o
�o 
5263 l   �n�m n     1    �l
�l 
2903 1    �k
�k 
sele�n  �m   �j	�i
�j 
5264	 l   
�h�g
 n     1    �f
�f 
2905 1    �e
�e 
sele�h  �g  �i   o      �d�d  0 selectionrange selectionRange� �c r    # m     � . R e a d i n g   b i b l i o g r a p h y . . . 1    "�b
�b 
1091�c  � m     �                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  �  l  % %�a�a   ( " look for a formatted bibliography    � D   l o o k   f o r   a   f o r m a t t e d   b i b l i o g r a p h y  r   % * m   % & �   n      1   ' )�`
�` 
txdl 1   & '�_
�_ 
ascr  !  O   +X"#" k   /W$$ %&% r   / 8'(' l  / 6)�^�]) I  / 6�\*�[
�\ .corecnte****       ***** n  / 2+,+ 2  0 2�Z
�Z 
w170, o   / 0�Y�Y 0 worddocument wordDocument�[  �^  �]  ( o      �X�X 0 
fieldcount 
fieldCount& -.- W   9 �/0/ k   A �11 232 r   A G454 l  A E6�W�V6 n   A E787 4   B E�U9
�U 
w1709 o   C D�T�T 0 
fieldcount 
fieldCount8 o   A B�S�S 0 worddocument wordDocument�W  �V  5 o      �R�R 0 currentfield currentField3 :;: r   H S<=< n   H Q>?> 1   M Q�Q
�Q 
1650? n   H M@A@ 1   I M�P
�P 
2469A o   H I�O�O 0 currentfield currentField= o      �N�N 0 	fieldcode 	fieldCode; BCB Z   T �DE�M�LD F   T nFGF =  T ]HIH n   T YJKJ 1   U Y�K
�K 
wFtPK o   T U�J�J 0 currentfield currentFieldI m   Y \�I
�I e183G QG E   ` jLML n   ` fNON 4   a f�HP
�H 
cworP m   d e�G�G O o   ` a�F�F 0 	fieldcode 	fieldCodeM m   f iQQ �RR  b i b l i o g r a p h yE k   q �SS TUT r   q �VWV c   q �XYX n   q �Z[Z 7  r ��E\]
�E 
cha \ m   x |�D�D ] m   } ��C�C��[ o   q r�B�B 0 	fieldcode 	fieldCodeY m   � ��A
�A 
TEXTW o      �@�@ ,0 bibliographysettings bibliographySettingsU ^�?^  S   � ��?  �M  �L  C _�>_ r   � �`a` \   � �bcb o   � ��=�= 0 
fieldcount 
fieldCountc m   � ��<�< a o      �;�; 0 
fieldcount 
fieldCount�>  0 =  = @ded o   = >�:�: 0 
fieldcount 
fieldCounte m   > ?�9�9  . f�8f Z   �Wgh�7�6g =  � �iji o   � ��5�5 0 
fieldcount 
fieldCountj m   � ��4�4  h l  �Sklmk k   �Snn opo l  � �qrsq r   � �tut n   � �vwv 1   � ��3
�3 
1717w 1   � ��2
�2 
seleu o      �1�1 0 
findobject 
findObjectr %  set up the find system in Word   s �xx >   s e t   u p   t h e   f i n d   s y s t e m   i n   W o r dp yzy r   � �{|{ m   � ��0
�0 boovtrue| n      }~} 1   � ��/
�/ 
1840~ o   � ��.�. 0 
findobject 
findObjectz � r   � ���� m   � ��-
�- e265� � n      ��� 1   � ��,
�, 
1862� o   � ��+�+ 0 
findobject 
findObject� ��� r   � ���� m   � ��� ��� " \ \ b i b l i o g r a p h y * \ }� n      ��� 1   � ��*
�* 
1650� o   � ��)�) 0 
findobject 
findObject� ��� I  � ��(��'
�( .miscslctnull���    obj � l  � ���&�%� I  � ��$��
�$ .sWRD1430null���     docu� o   � ��#�# 0 worddocument wordDocument� �"��
�" 
5263� m   � ��!�!  � � ��
�  
5264� m   � ���  �  �&  �%  �'  � ��� Q   ����� l  � ����� k   � ��� ��� r   � ���� m   � ���  � o      �� 0 	findcount 	findCount� ��� V   � ���� r   � ���� [   � ���� o   � ��� 0 	findcount 	findCount� m   � ��� � o      �� 0 	findcount 	findCount� l  � ����� I  � ����
� .sWRD1874null���     w124� o   � ��� 0 
findobject 
findObject�  �  �  �  � * $ count the number of instances found   � ��� H   c o u n t   t h e   n u m b e r   o f   i n s t a n c e s   f o u n d� R      ���
� .ascrerr ****      � ****�  �  � l  ����� k   ��� ��� I  ����
� .sysodlogaskr        TEXT� m   � ��� ��� � S o m e t h i n g   i s   w r o n g   w i t h   t h e   f i n d   s y s t e m   i n   M i c r o s o f t   W o r d ,   p l e a s e   q u i t   M i c r o s o f t   W o r d   t h e n   r e - o p e n� ���
� 
btns� J   ��� ��� m   � �� ���  O K�  � ���
� 
cbtn� J  
�� ��� m  �� ���  O K�  � �
��	
�
 
dflt� J  �� ��� m  �� ���  O K�  �	  � ��� L  ��  �  � H B this can break if advanced find and replace has been used in Word   � ��� �   t h i s   c a n   b r e a k   i f   a d v a n c e d   f i n d   a n d   r e p l a c e   h a s   b e e n   u s e d   i n   W o r d� ��� Z  S����� > ��� l ���� o  �� 0 	findcount 	findCount�  �  � m  � �   � l "C���� r  "C��� c  "A��� n  "=��� 7 -=����
�� 
cha � m  37���� � m  8<������� l "-������ c  "-��� n  ")��� 1  %)��
�� 
1650� 1  "%��
�� 
sele� m  ),��
�� 
TEXT��  ��  � m  =@��
�� 
TEXT� o      ���� ,0 bibliographysettings bibliographySettings� ' ! extract raw bibiography settings   � ��� B   e x t r a c t   r a w   b i b i o g r a p h y   s e t t i n g s� ��� = FI��� o  FG���� 0 	findcount 	findCount� m  GH����  � ���� l LO���� L  LO�� J  LN����  �   no bibliography found   � ��� ,   n o   b i b l i o g r a p h y   f o u n d��  �  �  l + % look for an unformatted bibliography   m ��� J   l o o k   f o r   a n   u n f o r m a t t e d   b i b l i o g r a p h y�7  �6  �8  # m   + ,���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ! ��� Q  YZ���� l \3���� k  \3�� ��� r  \���� J  \��� ��� m  \_�� ��� * < b i b D e s k D o c u m e n t N a m e =�    m  _b � 6 > < b i b l i o g r a p h y T e m p l a t e P a t h =  m  be � ( > < c i t e p T e m p l a t e P a t h = 	 m  eh

 � ( > < c i t e t T e m p l a t e P a t h =	  m  hk � $ > < p a r e n t h e s i s O p e n =  m  kn � & > < p a r e n t h e s i s C l o s e =  m  nq � * > < b i b l i o g r a p h y S o r t B y =  m  qt � ( > < b i b l i o g r a p h y S t y l e =  m  tw �  > < i t a l i c L a t i n =  !  m  wz"" �## 2 > < i t a l i c P u b l i c a t i o n T y p e s =! $%$ m  z}&& �'' 0 > < s u p e r s c r i p t R e f e r e n c e s =% ()( m  }�** �++ * > < n u m b e r e d R e f e r e n c e s =) ,-, m  ��.. �// $ > < n u m b e r S e p e r a t o r =- 0��0 m  ��11 �22  >��  � n     343 1  ����
�� 
txdl4 1  ����
�� 
ascr� 5��5 r  �3676 n  ��898 2 ����
�� 
citm9 o  ������ ,0 bibliographysettings bibliographySettings7 J      :: ;<; o      ���� *0 bibliographycommand bibliographyCommand< =>= o      ���� *0 bibdeskdocumentname bibDeskDocumentName> ?@? o      ���� 40 bibliographytemplatepath bibliographyTemplatePath@ ABA o      ���� &0 citeptemplatepath citepTemplatePathB CDC o      ���� &0 citettemplatepath citetTemplatePathD EFE o      ���� "0 parenthesisopen parenthesisOpenF GHG o      ���� $0 parenthesisclose parenthesisCloseH IJI o      ���� (0 bibliographysortby bibliographySortByJ KLK o      ���� &0 bibliographystyle bibliographyStyleL MNM o      ���� 0 italiclatin italicLatinN OPO o      ���� 00 italicpublicationtypes italicPublicationTypesP QRQ o      ���� .0 superscriptreferences superscriptReferencesR STS o      ���� (0 numberedreferences numberedReferencesT U��U o      ���� "0 numberseperator numberSeperator��  ��  � . ( extract bibliography settings as a list   � �VV P   e x t r a c t   b i b l i o g r a p h y   s e t t i n g s   a s   a   l i s t� R      ������
�� .ascrerr ****      � ****��  ��  � l ;ZWXYW k  ;ZZZ [\[ O  ;W]^] I ?V��_`
�� .sysodlogaskr        TEXT_ m  ?Baa �bb � T h e r e   i s   s o m e t h i n g   w r o n g   w i t h   t h e   b i b l i o g r a p h y   d a t a ,   p l e a s e   d e l e t e   t h e   b i b l i o g r a p h y   a n d   t r y   a g a i n` ��cd
�� 
btnsc J  EJee f��f m  EHgg �hh  O K��  d ��i��
�� 
dflti J  MRjj k��k m  MPll �mm  O K��  ��  ^ m  ;<nn�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  \ o��o L  XZ����  ��  X < 6 if anything is amiss, recreate the bibliography field   Y �pp l   i f   a n y t h i n g   i s   a m i s s ,   r e c r e a t e   t h e   b i b l i o g r a p h y   f i e l d� qrq l [bstus r  [bvwv m  [^xx �yy  ,w n     z{z 1  _a��
�� 
txdl{ 1  ^_��
�� 
ascrt "  recover lists from settings   u �|| 8   r e c o v e r   l i s t s   f r o m   s e t t i n g sr }~} r  cn� n  cj��� 2 fj��
�� 
citm� o  cf���� 0 italiclatin italicLatin� o      ���� 0 italiclatin italicLatin~ ��� r  oz��� n  ov��� 2 rv��
�� 
citm� o  or���� 00 italicpublicationtypes italicPublicationTypes� o      ���� 00 italicpublicationtypes italicPublicationTypes� ��� r  {���� m  {~�� ���  � n     ��� 1  ���
�� 
txdl� 1  ~��
�� 
ascr� ��� Q  ������ l ������ O  ����� k  ���� ��� e  ���� l �������� I �������
�� .coredoexnull���     ****� l �������� c  ����� 4  �����
�� 
psxf� l �������� c  ����� o  ������ 40 bibliographytemplatepath bibliographyTemplatePath� m  ����
�� 
TEXT��  ��  � m  ����
�� 
alis��  ��  ��  ��  ��  � ��� e  ���� l �������� I �������
�� .coredoexnull���     ****� l �������� c  ����� 4  �����
�� 
psxf� l �������� c  ����� o  ������ &0 citeptemplatepath citepTemplatePath� m  ����
�� 
TEXT��  ��  � m  ����
�� 
alis��  ��  ��  ��  ��  � ���� e  ���� l �������� I �������
�� .coredoexnull���     ****� l �������� c  ����� 4  �����
�� 
psxf� l �������� c  ����� o  ������ &0 citettemplatepath citetTemplatePath� m  ����
�� 
TEXT��  ��  � m  ����
�� 
alis��  ��  ��  ��  ��  ��  � m  �����                                                                                  MACS  alis    @  Macintosh HD                   BD ����
Finder.app                                                     ����            ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  � 7 1 check to make sure that the template files exist   � ��� b   c h e c k   t o   m a k e   s u r e   t h a t   t h e   t e m p l a t e   f i l e s   e x i s t� R      ������
�� .ascrerr ****      � ****��  ��  � l ������ k  ���� ��� O  ����� I ������
�� .sysodlogaskr        TEXT� m  ���� ��� � T e m p l a t e   f i l e s   n o t   f o u n d ,   p l e a s e   d e l e t e   t h e   b i b l i o g r a p h y   a n d   t r y   a g a i n� ����
�� 
btns� J  ���� ���� m  ���� ���  O K��  � �����
�� 
dflt� J  ���� ���� m  ���� ���  O K��  ��  � m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ���� L  ������  ��  � . ( if not, recreate the bibliography field   � ��� P   i f   n o t ,   r e c r e a t e   t h e   b i b l i o g r a p h y   f i e l d� ��� l ��������  � , & restore Word to its original position   � ��� L   r e s t o r e   W o r d   t o   i t s   o r i g i n a l   p o s i t i o n� ��� O  ���� k  ��� ��� Q  ����� I �������
�� .miscslctnull���    obj � o  ������  0 selectionrange selectionRange��  � R      ������
�� .ascrerr ****      � ****��  ��  � l ���� I �����
�� .miscslctnull���    obj � l ������ I ����
�� .sWRD1430null���     docu� o  ���� 0 worddocument wordDocument� ����
�� 
5263� l ������ n  ��� 1  ��
�� 
2905� 1  ��
�� 
sele��  ��  � �����
�� 
5264� l 	������ n  	��� 1  ��
�� 
2905� 1  	��
�� 
sele��  ��  ��  ��  ��  ��  � > 8 produces error if cursor was at the end of the document   � ��� p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n t� ��� r  ��� m  �� �   : F i n i s h e d   r e a d i n g   b i b l i o g r a p h y� 1  �~
�~ 
1091�  � m  ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � �} L   > J   =  o   !�|�| *0 bibdeskdocumentname bibDeskDocumentName  o  !"�{�{ 40 bibliographytemplatepath bibliographyTemplatePath 	
	 o  "#�z�z &0 citeptemplatepath citepTemplatePath
  o  #$�y�y &0 citettemplatepath citetTemplatePath  o  $%�x�x "0 parenthesisopen parenthesisOpen  o  %&�w�w $0 parenthesisclose parenthesisClose  o  &'�v�v (0 bibliographysortby bibliographySortBy  o  '*�u�u &0 bibliographystyle bibliographyStyle  o  *-�t�t 0 italiclatin italicLatin  o  -0�s�s 00 italicpublicationtypes italicPublicationTypes  o  03�r�r .0 superscriptreferences superscriptReferences  o  36�q�q (0 numberedreferences numberedReferences �p o  69�o�o "0 numberseperator numberSeperator�p  �}  �  l     �n�m�l�n  �m  �l    !  l     �k"#�k  "  -   # �$$  -! %&% l     �j�i�h�j  �i  �h  & '(' i    )*) I      �g+�f�g 60 writebibliographysettings writeBibliographySettings+ ,�e, o      �d�d ,0 bibliographysettings bibliographySettings�e  �f  * k    �-- ./. l     �c01�c  0   global setup   1 �22    g l o b a l   s e t u p/ 343 O     $565 k    #77 898 r    	:;: 1    �b
�b 
1003; o      �a�a 0 worddocument wordDocument9 <=< r   
 >?> I  
 �`@A
�` .sWRD1430null���     docu@ o   
 �_�_ 0 worddocument wordDocumentA �^BC
�^ 
5263B l   D�]�\D n    EFE 1    �[
�[ 
2903F 1    �Z
�Z 
sele�]  �\  C �YG�X
�Y 
5264G l   H�W�VH n    IJI 1    �U
�U 
2905J 1    �T
�T 
sele�W  �V  �X  ? o      �S�S  0 selectionrange selectionRange= K�RK r    #LML m    NN �OO . W r i t i n g   b i b l i o g r a p h y . . .M 1    "�Q
�Q 
1091�R  6 m     PP�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  4 QRQ O   % 0STS r   ) /UVU 4  ) -�PW
�P 
docuW m   + ,�O�O V o      �N�N "0 bibdeskdocument bibDeskDocumentT m   % &XX�                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  R YZY l  1 1�M[\�M  [ "  parse bibliography settings   \ �]] 8   p a r s e   b i b l i o g r a p h y   s e t t i n g sZ ^_^ r   1 6`a` m   1 2bb �cc  a n     ded 1   3 5�L
�L 
txdle 1   2 3�K
�K 
ascr_ fgf r   7 �hih o   7 8�J�J ,0 bibliographysettings bibliographySettingsi J      jj klk o      �I�I *0 bibdeskdocumentname bibDeskDocumentNamel mnm o      �H�H 40 bibliographytemplatepath bibliographyTemplatePathn opo o      �G�G &0 citeptemplatepath citepTemplatePathp qrq o      �F�F &0 citettemplatepath citetTemplatePathr sts o      �E�E "0 parenthesisopen parenthesisOpent uvu o      �D�D $0 parenthesisclose parenthesisClosev wxw o      �C�C (0 bibliographysortby bibliographySortByx yzy o      �B�B &0 bibliographystyle bibliographyStylez {|{ o      �A�A 0 italiclatin italicLatin| }~} o      �@�@ 00 italicpublicationtypes italicPublicationTypes~ � o      �?�? .0 superscriptreferences superscriptReferences� ��� o      �>�> (0 numberedreferences numberedReferences� ��=� o      �<�< "0 numberseperator numberSeperator�=  g ��� l  � ��;���;  � 8 2 parse custom lists from the bibliography settings   � ��� d   p a r s e   c u s t o m   l i s t s   f r o m   t h e   b i b l i o g r a p h y   s e t t i n g s� ��� r   � ���� m   � ��� ���  ,� n     ��� 1   � ��:
�: 
txdl� 1   � ��9
�9 
ascr� ��� r   � ���� n   � ���� 2  � ��8
�8 
citm� o   � ��7�7 0 italiclatin italicLatin� o      �6�6 0 italiclatin italicLatin� ��� r   � ���� n   � ���� 2  � ��5
�5 
citm� o   � ��4�4 00 italicpublicationtypes italicPublicationTypes� o      �3�3 00 italicpublicationtypes italicPublicationTypes� ��� r   � ���� m   � ��� ���  � n     ��� 1   � ��2
�2 
txdl� 1   � ��1
�1 
ascr� ��� l  � ��0���0  � &   look for formatted bibliography   � ��� @   l o o k   f o r   f o r m a t t e d   b i b l i o g r a p h y� ��� O   �N��� k   �M�� ��� r   � ���� l  � ���/�.� I  � ��-��,
�- .corecnte****       ****� n  � ���� 2  � ��+
�+ 
w170� o   � ��*�* 0 worddocument wordDocument�,  �/  �.  � o      �)�) 0 
fieldcount 
fieldCount� ��(� Z   �M���'�&� >  � ���� o   � ��%�% 0 
fieldcount 
fieldCount� m   � ��$�$  � W   �I��� k  D�� ��� Z  :���#�"� E  ��� n  ��� 1  �!
�! 
1650� n  ��� 1  � 
�  
2469� n  ��� 4  ��
� 
w170� o  �� 0 
fieldcount 
fieldCount� o  �� 0 worddocument wordDocument� m  �� ��� $ A D D I N   b i b l i o g r a p h y� k  6�� ��� r  *��� n  &��� 4  &��
� 
w170� o  "%�� 0 
fieldcount 
fieldCount� o  �� 0 worddocument wordDocument� o      �� &0 bibliographyfield bibliographyField� ��� r  +4��� m  +,�
� boovtrue� n      ��� 1  /3�
� 
2482� o  ,/�� &0 bibliographyfield bibliographyField� ���  S  56�  �#  �"  � ��� r  ;D��� \  ;@��� o  ;>�� 0 
fieldcount 
fieldCount� m  >?�� � o      �� 0 
fieldcount 
fieldCount�  � = ��� o  �� 0 
fieldcount 
fieldCount� m  ��  �'  �&  �(  � m   � ����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� l OO����  � 3 - if none found, create the bibliography field   � ��� Z   i f   n o n e   f o u n d ,   c r e a t e   t h e   b i b l i o g r a p h y   f i e l d� ��� Z  O������ = OT��� o  OR�� 0 
fieldcount 
fieldCount� m  RS�
�
  � k  W��� ��� Q  W����� O  Z���� k  ^��� ��� I ^e�	��
�	 .sysodlogaskr        TEXT� m  ^a�� ��� j N o   b i b l i o g r a p h y   f o u n d ,   d o   y o u   w i s h   t o   c r e a t e   o n e   n o w ?�  � ��� r  f�� � I f���
� .corecrel****      � null�   �
� 
kocl m  jm�
� 
w170 �
� 
insh n  pv  ;  uv n  pu	 1  qu�
� 
wTxR	 o  pq� �  0 worddocument wordDocument ��
��
�� 
prdt
 K  y� ��
�� 
wFtP m  |��
�� e183G # ��
�� 
2482 m  ����
�� boovtrue ����
�� 
2477 m  �� �  *��  ��    o      ���� &0 bibliographyfield bibliographyField�  � m  Z[�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � R      ������
�� .ascrerr ****      � ****��  ��  � l �� L  �� J  ������   - ' return the bibliography settings empty    � N   r e t u r n   t h e   b i b l i o g r a p h y   s e t t i n g s   e m p t y�  Z  ������ = �� o  ������ *0 bibdeskdocumentname bibDeskDocumentName m  �� �     l ��!"#! O  ��$%$ r  ��&'& n  ��()( 1  ����
�� 
pnam) o  ������ "0 bibdeskdocument bibDeskDocument' o      ���� *0 bibdeskdocumentname bibDeskDocumentName% m  ��**�                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  " = 7 if default settings were used, set the BibDesk library   # �++ n   i f   d e f a u l t   s e t t i n g s   w e r e   u s e d ,   s e t   t h e   B i b D e s k   l i b r a r y��  ��   ,-, Z  ��./����. = ��010 o  ������ 40 bibliographytemplatepath bibliographyTemplatePath1 m  ��22 �33  / l ��4564 O  ��787 r  ��9:9 n  ��;<; 1  ����
�� 
psxp< l ��=����= I ������>
�� .sysostdfalis    ��� null��  > ��?@
�� 
prmp? m  ��AA �BB \ P l e a s e   c h o o s e   W o r d S c r i p t   T e m p l a t e   B i b l i o g r a p h y@ ��C��
�� 
ftypC J  ��DD EFE m  ��GG �HH  p u b l i c . r t fF I��I m  ��JJ �KK , c o m . m i c r o s o f t . w o r d . d o c��  ��  ��  ��  : o      ���� 40 bibliographytemplatepath bibliographyTemplatePath8 m  ��LL�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  5 < 6 if default settings were used, set the template paths   6 �MM l   i f   d e f a u l t   s e t t i n g s   w e r e   u s e d ,   s e t   t h e   t e m p l a t e   p a t h s��  ��  - NON l ��PQRP r  ��STS m  ��UU �VV  /T n     WXW 1  ����
�� 
txdlX 1  ����
�� 
ascrQ &    guess the other template paths   R �YY @     g u e s s   t h e   o t h e r   t e m p l a t e   p a t h sO Z[Z r  �
\]\ c  �^_^ l �`����` n  �aba 7���cd
�� 
citmc m  ������ d m  �������b o  ������ 40 bibliographytemplatepath bibliographyTemplatePath��  ��  _ m  ��
�� 
TEXT] o      ���� *0 assumedtemplatepath assumedTemplatePath[ efe r  ghg m  ii �jj  h n     klk 1  ��
�� 
txdll 1  ��
�� 
ascrf mnm O  copo Q  bqr��q k  Yss tut r  :vwv n  8xyx 1  48��
�� 
psxpy l 4z����z c  4{|{ 4  0��}
�� 
psxf} l  /~����~ c   /� b   +��� b   '��� o   #���� *0 assumedtemplatepath assumedTemplatePath� m  #&�� ���  /� m  '*�� ��� V W o r d S c r i p t   T e m p l a t e   A u t h o r - Y e a r   ( c i t e p ) . t x t� m  +.��
�� 
TEXT��  ��  | m  03��
�� 
alis��  ��  w o      ���� &0 citeptemplatepath citepTemplatePathu ���� r  ;Y��� n  ;W��� 1  SW��
�� 
psxp� l ;S������ c  ;S��� 4  ;O���
�� 
psxf� l ?N������ c  ?N��� b  ?J��� b  ?F��� o  ?B���� *0 assumedtemplatepath assumedTemplatePath� m  BE�� ���  /� m  FI�� ��� R W o r d S c r i p t   T e m p l a t e   Y e a r   O n l y   ( c i t e t ) . t x t� m  JM��
�� 
TEXT��  ��  � m  OR��
�� 
alis��  ��  � o      ���� &0 citettemplatepath citetTemplatePath��  r R      ������
�� .ascrerr ****      � ****��  ��  ��  p m  ���                                                                                  MACS  alis    @  Macintosh HD                   BD ����
Finder.app                                                     ����            ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  n ��� Z  d�������� = di��� o  de���� &0 citeptemplatepath citepTemplatePath� m  eh�� ���  � l l����� O  l���� r  p���� n  p���� 1  ����
�� 
psxp� l p������� I p������
�� .sysostdfalis    ��� null��  � ����
�� 
prmp� m  tw�� ��� j P l e a s e   c h o o s e   W o r d S c r i p t   T e m p l a t e   A u t h o r - Y e a r   ( c i t e p )� �����
�� 
ftyp� J  z�� ���� m  z}�� ��� " p u b l i c . p l a i n - t e x t��  ��  ��  ��  � o      ���� &0 citeptemplatepath citepTemplatePath� m  lm���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � !   otherwise choose manually   � ��� 6     o t h e r w i s e   c h o o s e   m a n u a l l y��  ��  � ���� Z  ��������� = ����� o  ������ &0 citettemplatepath citetTemplatePath� m  ���� ���  � O  ����� r  ����� n  ����� 1  ����
�� 
psxp� l �������� I �������
�� .sysostdfalis    ��� null��  � ����
�� 
prmp� m  ���� ��� f P l e a s e   c h o o s e   W o r d S c r i p t   T e m p l a t e   Y e a r   O n l y   ( c i t e t )� �����
�� 
ftyp� J  ���� ���� m  ���� ��� " p u b l i c . p l a i n - t e x t��  ��  ��  ��  � o      ���� &0 citettemplatepath citetTemplatePath� m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ��  ��  ��  �  �  � ��� l ��������  � * $ write the bibiography field in Word   � ��� H   w r i t e   t h e   b i b i o g r a p h y   f i e l d   i n   W o r d� ��� O  �\��� k  �[�� ��� r  ����� m  ���� ���  ,� n     ��� 1  ����
�� 
txdl� 1  ����
�� 
ascr� ��� r  �-��� b  �!��� b  ���� b  ���� b  ���� b  ���� b  ���� b  ���� b  �	��� b  ���� b  ���� b  �� � b  �� b  �� b  �� b  �� b  ��	
	 b  �� b  �� b  �� b  �� b  �� b  �� b  �� b  �� b  �� b  �� b  ��  m  ��!! �"" (   A D D I N   b i b l i o g r a p h y {  m  ��## �$$ * < b i b D e s k D o c u m e n t N a m e = o  ������ *0 bibdeskdocumentname bibDeskDocumentName m  ��%% �&& 6 > < b i b l i o g r a p h y T e m p l a t e P a t h = o  ������ 40 bibliographytemplatepath bibliographyTemplatePath m  ��'' �(( ( > < c i t e p T e m p l a t e P a t h = o  ������ &0 citeptemplatepath citepTemplatePath m  ��)) �** ( > < c i t e t T e m p l a t e P a t h = o  ������ &0 citettemplatepath citetTemplatePath m  ��++ �,, $ > < p a r e n t h e s i s O p e n = o  ������ "0 parenthesisopen parenthesisOpen m  ��-- �.. & > < p a r e n t h e s i s C l o s e =
 o  ������ $0 parenthesisclose parenthesisClose m  ��// �00 * > < b i b l i o g r a p h y S o r t B y = o  ������ (0 bibliographysortby bibliographySortBy m  ��11 �22 ( > < b i b l i o g r a p h y S t y l e = o  ������ &0 bibliographystyle bibliographyStyle  m  � 33 �44  > < i t a l i c L a t i n =� o  ���� 0 italiclatin italicLatin� m  55 �66 2 > < i t a l i c P u b l i c a t i o n T y p e s =� o  ���� 00 italicpublicationtypes italicPublicationTypes� m  	77 �88 0 > < s u p e r s c r i p t R e f e r e n c e s =� o  ���� .0 superscriptreferences superscriptReferences� m  99 �:: * > < n u m b e r e d R e f e r e n c e s =� o  ���� (0 numberedreferences numberedReferences� m  ;; �<< $ > < n u m b e r S e p e r a t o r =� o  ���� "0 numberseperator numberSeperator� m   == �>>  > }� n      ?@? 1  (,��
�� 
1650@ n  !(ABA 1  $(��
�� 
2469B o  !$���� &0 bibliographyfield bibliographyField� CDC r  .5EFE m  .1GG �HH  F n     IJI 1  24��
�� 
txdlJ 1  12��
�� 
ascrD K��K Q  6[LMNL l 9>OPQO I 9>��R��
�� .miscslctnull���    obj R o  9:����  0 selectionrange selectionRange��  P 2 , restore the cursor to its original position   Q �SS X   r e s t o r e   t h e   c u r s o r   t o   i t s   o r i g i n a l   p o s i t i o nM R      ������
�� .ascrerr ****      � ****��  ��  N l F[TUVT I F[��W��
�� .miscslctnull���    obj W l FWX����X I FW�YZ
� .sWRD1430null���     docuY o  FG�~�~ 0 worddocument wordDocumentZ �}[\
�} 
5263[ l HM]�|�{] n  HM^_^ 1  KM�z
�z 
2905_ 1  HK�y
�y 
sele�|  �{  \ �x`�w
�x 
5264` l NSa�v�ua n  NSbcb 1  QS�t
�t 
2905c 1  NQ�s
�s 
sele�v  �u  �w  ��  ��  ��  U > 8 produces error if cursor was at the end of the document   V �dd p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n t��  � m  ��ee�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � fgf l ]]�rhi�r  h , & restore Word to its original position   i �jj L   r e s t o r e   W o r d   t o   i t s   o r i g i n a l   p o s i t i o ng klk O  ]�mnm k  a�oo pqp Q  a�rstr I di�qu�p
�q .miscslctnull���    obj u o  de�o�o  0 selectionrange selectionRange�p  s R      �n�m�l
�n .ascrerr ****      � ****�m  �l  t l q�vwxv I q��ky�j
�k .miscslctnull���    obj y l q�z�i�hz I q��g{|
�g .sWRD1430null���     docu{ o  qr�f�f 0 worddocument wordDocument| �e}~
�e 
5263} l sx�d�c n  sx��� 1  vx�b
�b 
2905� 1  sv�a
�a 
sele�d  �c  ~ �`��_
�` 
5264� l y~��^�]� n  y~��� 1  |~�\
�\ 
2905� 1  y|�[
�[ 
sele�^  �]  �_  �i  �h  �j  w > 8 produces error if cursor was at the end of the document   x ��� p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n tq ��Z� r  ����� m  ���� ��� : F i n i s h e d   w r i t i n g   b i b l i o g r a p h y� 1  ���Y
�Y 
1091�Z  n m  ]^���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  l ��X� L  ���� J  ���� ��� o  ���W�W *0 bibdeskdocumentname bibDeskDocumentName� ��� o  ���V�V 40 bibliographytemplatepath bibliographyTemplatePath� ��� o  ���U�U &0 citeptemplatepath citepTemplatePath� ��� o  ���T�T &0 citettemplatepath citetTemplatePath� ��� o  ���S�S "0 parenthesisopen parenthesisOpen� ��� o  ���R�R $0 parenthesisclose parenthesisClose� ��� o  ���Q�Q (0 bibliographysortby bibliographySortBy� ��� o  ���P�P &0 bibliographystyle bibliographyStyle� ��� o  ���O�O 0 italiclatin italicLatin� ��� o  ���N�N 00 italicpublicationtypes italicPublicationTypes� ��� o  ���M�M .0 superscriptreferences superscriptReferences� ��� o  ���L�L (0 numberedreferences numberedReferences� ��K� o  ���J�J "0 numberseperator numberSeperator�K  �X  ( ��� l     �I�H�G�I  �H  �G  � ��� l     �F���F  �  -   � ���  -� ��� l     �E�D�C�E  �D  �C  � ��� i    ��� I      �B��A�B $0 formatreferences formatReferences� ��@� o      �?�? ,0 bibliographysettings bibliographySettings�@  �A  � k    	��� ��� l     �>���>  �   global setup   � ���    g l o b a l   s e t u p� ��� O     $��� k    #�� ��� r    	��� 1    �=
�= 
1003� o      �<�< 0 worddocument wordDocument� ��� r   
 ��� I  
 �;��
�; .sWRD1430null���     docu� o   
 �:�: 0 worddocument wordDocument� �9��
�9 
5263� l   ��8�7� n    ��� 1    �6
�6 
2903� 1    �5
�5 
sele�8  �7  � �4��3
�4 
5264� l   ��2�1� n    ��� 1    �0
�0 
2905� 1    �/
�/ 
sele�2  �1  �3  � o      �.�.  0 selectionrange selectionRange� ��-� r    #��� m    �� ��� 0 F o r m a t t i n g   r e f e r e n c e s . . .� 1    "�,
�, 
1091�-  � m     ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� O   % 0��� r   ) /��� 4  ) -�+�
�+ 
docu� m   + ,�*�* � o      �)�) "0 bibdeskdocument bibDeskDocument� m   % &���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  � ��� l  1 1�(���(  � "  parse bibliography settings   � ��� 8   p a r s e   b i b l i o g r a p h y   s e t t i n g s� ��� r   1 6��� m   1 2�� ���  � n     ��� 1   3 5�'
�' 
txdl� 1   2 3�&
�& 
ascr� ��� r   7 ���� o   7 8�%�% ,0 bibliographysettings bibliographySettings� J      �� ��� o      �$�$ *0 bibdeskdocumentname bibDeskDocumentName� ��� o      �#�# 40 bibliographytemplatepath bibliographyTemplatePath� ��� o      �"�" &0 citeptemplatepath citepTemplatePath� ��� o      �!�! &0 citettemplatepath citetTemplatePath� ��� o      � �  "0 parenthesisopen parenthesisOpen�    o      �� $0 parenthesisclose parenthesisClose  o      �� (0 bibliographysortby bibliographySortBy  o      �� &0 bibliographystyle bibliographyStyle  o      �� 0 italiclatin italicLatin 	 o      �� 00 italicpublicationtypes italicPublicationTypes	 

 o      �� .0 superscriptreferences superscriptReferences  o      �� (0 numberedreferences numberedReferences � o      �� "0 numberseperator numberSeperator�  �  l  � ���   8 2 parse custom lists from the bibliography settings    � d   p a r s e   c u s t o m   l i s t s   f r o m   t h e   b i b l i o g r a p h y   s e t t i n g s  r   � � m   � � �  , n      1   � ��
� 
txdl 1   � ��
� 
ascr  r   � � n   � � !  2  � ��
� 
citm! o   � ��� 0 italiclatin italicLatin o      �� 0 italiclatin italicLatin "#" r   � �$%$ n   � �&'& 2  � ��
� 
citm' o   � ��� 00 italicpublicationtypes italicPublicationTypes% o      �� 00 italicpublicationtypes italicPublicationTypes# ()( r   � �*+* m   � �,, �--  + n     ./. 1   � ��
� 
txdl/ 1   � ��
� 
ascr) 010 l  � ��23�  2 _ Y read the plain text template data (should this move to 'format author-year references'?)   3 �44 �   r e a d   t h e   p l a i n   t e x t   t e m p l a t e   d a t a   ( s h o u l d   t h i s   m o v e   t o   ' f o r m a t   a u t h o r - y e a r   r e f e r e n c e s ' ? )1 565 r   � �787 I  � ��
9�	
�
 .rdwrread****        ****9 o   � ��� &0 citeptemplatepath citepTemplatePath�	  8 o      �� 0 citeptemplate citepTemplate6 :;: r   � �<=< I  � ��>�
� .rdwrread****        ****> o   � ��� &0 citettemplatepath citetTemplatePath�  = o      �� 0 citettemplate citetTemplate; ?@? l  � ��AB�  A !  intitialise document lists   B �CC 6   i n t i t i a l i s e   d o c u m e n t   l i s t s@ DED r   � �FGF J   � ���  G o      � �  0 citekeylist citeKeyListE HIH r   �JKJ J   � ����  K o      ���� "0 publicationlist publicationListI LML r  NON J  ����  O o      ���� "0 ibidcitekeylist ibidCiteKeyListM PQP r  RSR J  ����  S o      ���� $0 italicreferences italicReferencesQ TUT l ��VW��  V &   begin formatting the references   W �XX @   b e g i n   f o r m a t t i n g   t h e   r e f e r e n c e sU YZY O  	�[\[ k  	�]] ^_^ I ������
�� .miscactvnull��� ��� null��  ��  _ `a` l ��bc��  b !  establish the format range   c �dd 6   e s t a b l i s h   t h e   f o r m a t   r a n g ea efe Z  fgh��ig = (jkj l "l����l n  "mnm 1   "��
�� 
2903n 1   ��
�� 
sele��  ��  k l "'o����o n  "'pqp 1  %'��
�� 
2905q 1  "%��
�� 
sele��  ��  h Z  +^rs��tr = +6uvu n  +2wxw 1  .2��
�� 
1650x 1  +.��
�� 
selev m  25yy �zz  ]s l 9H{|}{ r  9H~~ n  9D��� 2 @D��
�� 
cwor� n  9@��� 2 <@��
�� 
csen� 1  9<��
�� 
sele o      ���� 0 formatrange formatRange|   just edited page number   } ��� 0   j u s t   e d i t e d   p a g e   n u m b e r��  t k  K^�� ��� l KP���� r  KP��� o  KL���� 0 worddocument wordDocument� o      ���� 0 formatrange formatRange�   format entire document   � ��� .   f o r m a t   e n t i r e   d o c u m e n t� ���� I Q^�����
�� .miscslctnull���    obj � l QZ������ I QZ����
�� .sWRD1430null���     docu� o  QR���� 0 worddocument wordDocument� ����
�� 
5263� m  ST����  � �����
�� 
5264� m  UV����  ��  ��  ��  ��  ��  ��  i l af���� r  af��� o  ab����  0 selectionrange selectionRange� o      ���� 0 formatrange formatRange�   format selection only   � ��� ,   f o r m a t   s e l e c t i o n   o n l yf ��� l gg������  � %  set up the find system in Word   � ��� >   s e t   u p   t h e   f i n d   s y s t e m   i n   W o r d� ��� r  gr��� n  gn��� 1  jn��
�� 
1717� 1  gj��
�� 
sele� o      ���� 0 
findobject 
findObject� ��� r  s|��� m  st��
�� boovtrue� n      ��� 1  w{��
�� 
1840� o  tw���� 0 
findobject 
findObject� ��� r  }���� m  }���
�� e265� � n      ��� 1  ����
�� 
1862� o  ������ 0 
findobject 
findObject� ��� l ��������  � 8 2 convert unformatted cite commands to ADDIN fields   � ��� d   c o n v e r t   u n f o r m a t t e d   c i t e   c o m m a n d s   t o   A D D I N   f i e l d s� ��� X  ������� k  ���� ��� r  ����� b  ����� b  ����� m  ���� ���  \ \� o  ������ 0 citecommand citeCommand� m  ���� ���  * \ }� n      ��� 1  ����
�� 
1650� o  ������ 0 
findobject 
findObject� ��� I �������
�� .miscslctnull���    obj � o  ������ 0 formatrange formatRange��  � ���� Q  ������ V  �b��� k  �]�� ��� r  ����� n  ����� 1  ����
�� 
1650� 1  ����
�� 
sele� o      ���� 0 citecommand citeCommand� ��� r  ����� m  ���� ���  � n      ��� 1  ����
�� 
1650� 1  ����
�� 
sele� ��� I � �����
�� .corecrel****      � null��  � ����
�� 
kocl� m  ����
�� 
w170� ����
�� 
insh� l ������� n  ���� 1  ���
�� 
wTxR� 1  ����
�� 
sele��  ��  � �����
�� 
prdt� K  �� ����
�� 
wFtP� m  	��
�� e183G #� ����
�� 
2482� m  ��
�� boovtrue� �����
�� 
2477� m  �� ���  *��  ��  � ��� r  !=��� n  !9��� 4 49���
�� 
w170� m  78���� � l !4������ I !4����
�� .sWRD1430null���     docu� o  !"���� 0 worddocument wordDocument� ��	 	
�� 
5263	  l #*	����	 \  #*			 l #(	����	 n  #(			 1  &(��
�� 
2903	 1  #&��
�� 
sele��  ��  	 m  ()���� ��  ��  	 ��	��
�� 
5264	 l +0		����		 n  +0	
		
 1  .0��
�� 
2905	 1  +.��
�� 
sele��  ��  ��  ��  ��  � o      ���� 0 currentfield currentField� 	��	 r  >]			 b  >Q			 m  >A		 �		    A D D I N  	 n  AP			 7 DP��		
�� 
cha 	 m  JL���� 	 m  MO������	 o  AD���� 0 citecommand citeCommand	 n      			 1  X\��
�� 
1650	 n  QX			 1  TX��
�� 
2469	 o  QT���� 0 currentfield currentField��  � l ��	����	 I ����	��
�� .sWRD1874null���     w124	 o  ������ 0 
findobject 
findObject��  ��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � l j�				 k  j�	 	  	!	"	! I j���	#	$
�� .sysodlogaskr        TEXT	# m  jm	%	% �	&	& � S o m e t h i n g   i s   w r o n g   w i t h   t h e   f i n d   s y s t e m   i n   M i c r o s o f t   W o r d ,   p l e a s e   q u i t   M i c r o s o f t   W o r d   t h e n   r e - o p e n	$ ��	'	(
�� 
btns	' J  pu	)	) 	*��	* m  ps	+	+ �	,	,  O K��  	( ��	-��
�� 
dflt	- J  x}	.	. 	/��	/ m  x{	0	0 �	1	1  O K��  ��  	" 	2��	2 L  ����  ��  	 H B this can break if advanced find and replace has been used in Word   	 �	3	3 �   t h i s   c a n   b r e a k   i f   a d v a n c e d   f i n d   a n d   r e p l a c e   h a s   b e e n   u s e d   i n   W o r d��  �� 0 citecommand citeCommand� J  ��	4	4 	5	6	5 m  ��	7	7 �	8	8 
 c i t e p	6 	9	:	9 m  ��	;	; �	<	< 
 c i t e t	: 	=	>	= m  ��	?	? �	@	@  c i t e a l p	> 	A	B	A m  ��	C	C �	D	D  c i t e a l t	B 	E�~	E m  ��	F	F �	G	G  b i b l i o g r a p h y�~  � 	H	I	H Y  ��	J�}	K	L�|	J l ��	M	N	O	M k  ��	P	P 	Q	R	Q r  ��	S	T	S n  ��	U	V	U 4  ���{	W
�{ 
w170	W o  ���z�z 0 fieldnumber fieldNumber	V o  ���y�y 0 formatrange formatRange	T o      �x�x 0 currentfield currentField	R 	X	Y	X r  ��	Z	[	Z n  ��	\	]	\ 1  ���w
�w 
1650	] n  ��	^	_	^ 1  ���v
�v 
2469	_ o  ���u�u 0 currentfield currentField	[ o      �t�t 0 	fieldcode 	fieldCode	Y 	`�s	` Z  ��	a	b�r�q	a F  ��	c	d	c = ��	e	f	e n  ��	g	h	g 1  ���p
�p 
wFtP	h o  ���o�o 0 currentfield currentField	f m  ���n
�n e183G Q	d E  ��	i	j	i J  ��	k	k 	l	m	l m  ��	n	n �	o	o 
 c i t e p	m 	p	q	p m  ��	r	r �	s	s 
 c i t e t	q 	t	u	t m  ��	v	v �	w	w  c i t e a l p	u 	x�m	x m  ��	y	y �	z	z  c i t e a l t�m  	j n  ��	{	|	{ 4  ���l	}
�l 
cwor	} m  ���k�k 	| o  ���j�j 0 	fieldcode 	fieldCode	b l ��	~		�	~ k  ��	�	� 	�	�	� I ���i	��h
�i .miscslctnull���    obj 	� o  ���g�g 0 currentfield currentField�h  	� 	�	�	� l ��	�	�	�	� r  ��	�	�	� J  ��	�	� 	�	�	� m  ��	�	� �	�	�  {	� 	��f	� m  ��	�	� �	�	�  }�f  	� n     	�	�	� 1  ���e
�e 
txdl	� 1  ���d
�d 
ascr	�   parse cite key   	� �	�	�    p a r s e   c i t e   k e y	� 	�	�	� r  �
	�	�	� n  �	�	�	� 4  �c	�
�c 
citm	� m  �b�b 	� o  ��a�a 0 	fieldcode 	fieldCode	� o      �`�` 0 citekey citeKey	� 	�	�	� l 	�	�	�	� r  	�	�	� J  	�	� 	�	�	� m  	�	� �	�	�  [	� 	��_	� m  	�	� �	�	�  ]�_  	� n     	�	�	� 1  �^
�^ 
txdl	� 1  �]
�] 
ascr	�   parse prefix/suffix   	� �	�	� (   p a r s e   p r e f i x / s u f f i x	� 	�	�	� r  	�	�	� m  	�	� �	�	�  	� o      �\�\ "0 referenceprefix referencePrefix	� 	�	�	� r   '	�	�	� m   #	�	� �	�	�  	� o      �[�[ "0 referencesuffix referenceSuffix	� 	�	�	� Z  (v	�	�	��Z	� = (7	�	�	� l (3	��Y�X	� I (3�W	��V
�W .corecnte****       ****	� n  (/	�	�	� 2 +/�U
�U 
citm	� o  (+�T�T 0 	fieldcode 	fieldCode�V  �Y  �X  	� m  36�S�S 	� l :S	�	�	�	� k  :S	�	� 	�	�	� r  :F	�	�	� n  :B	�	�	� 4  =B�R	�
�R 
citm	� m  @A�Q�Q 	� o  :=�P�P 0 	fieldcode 	fieldCode	� o      �O�O "0 referenceprefix referencePrefix	� 	��N	� r  GS	�	�	� n  GO	�	�	� 4  JO�M	�
�M 
citm	� m  MN�L�L 	� o  GJ�K�K 0 	fieldcode 	fieldCode	� o      �J�J "0 referencesuffix referenceSuffix�N  	� #  two pairs of square brackets   	� �	�	� :   t w o   p a i r s   o f   s q u a r e   b r a c k e t s	� 	�	�	� = Vc	�	�	� l Va	��I�H	� I Va�G	��F
�G .corecnte****       ****	� n  V]	�	�	� 2 Y]�E
�E 
citm	� o  VY�D�D 0 	fieldcode 	fieldCode�F  �I  �H  	� m  ab�C�C 	� 	��B	� l fr	�	�	�	� r  fr	�	�	� n  fn	�	�	� 4  in�A	�
�A 
citm	� m  lm�@�@ 	� o  fi�?�? 0 	fieldcode 	fieldCode	� o      �>�> "0 referencesuffix referenceSuffix	� "  one pair of square brackets   	� �	�	� 8   o n e   p a i r   o f   s q u a r e   b r a c k e t s�B  �Z  	� 	�	�	� O  wQ	�	�	� l {P	�	�	�	� k  {P	�	� 	�	�	� l {�	�	�	�	� r  {�	�	�	� J  {}�=�=  	� o      �<�< ,0 multiplepublications multiplePublications	� / ) intitialise the multiple reference lists   	� �	�	� R   i n t i t i a l i s e   t h e   m u l t i p l e   r e f e r e n c e   l i s t s	� 	�	�	� r  ��	�	�	� J  ���;�;  	� o      �:�: $0 multiplecitekeys multipleCiteKeys	� 	�
 	� l ��



 r  ��


 m  ��

 �

  ,
 n     

	
 1  ���9
�9 
txdl
	 1  ���8
�8 
ascr
 @ : multiple cite keys must be seperated by commas in BibDesk   
 �



 t   m u l t i p l e   c i t e   k e y s   m u s t   b e   s e p e r a t e d   b y   c o m m a s   i n   B i b D e s k
  
�7
 Y  �P
�6

�5
 l �K



 k  �K

 


 r  ��


 n  ��


 4  ���4

�4 
cobj
 o  ���3�3 0 eachinteger eachInteger
 n  ��


 2 ���2
�2 
citm
 o  ���1�1 0 citekey citeKey
 o      �0�0  0 currentcitekey currentCiteKey
 


 r  ��


 l ��
 �/�.
  6 ��
!
"
! n ��
#
$
# 2 ���-
�- 
bibi
$ o  ���,�, "0 bibdeskdocument bibDeskDocument
" = ��
%
&
% 1  ���+
�+ 
ckey
& o  ���*�*  0 currentcitekey currentCiteKey�/  �.  
 o      �)�) (0 currentpublication currentPublication
 
'
(
' Z  �4
)
*�(�'
) = ��
+
,
+ l ��
-�&�%
- I ���$
.�#
�$ .corecnte****       ****
. o  ���"�" (0 currentpublication currentPublication�#  �&  �%  
, m  ���!�!  
* l �0
/
0
1
/ O  �0
2
3
2 k  �/
4
4 
5
6
5 r  ��
7
8
7 m  ��� 
�  boovtrue
8 n      
9
:
9 1  ���
� 
2482
: o  ���� 0 currentfield currentField
6 
;
<
; I ���
=�
� .miscslctnull���    obj 
= n  ��
>
?
> 1  ���
� 
2469
? o  ���� 0 currentfield currentField�  
< 
@
A
@ I ��
B�
� .miscslctnull���    obj 
B l �
C��
C I ��
D
E
� .sWRD1430null���     docu
D o  ���� 0 worddocument wordDocument
E �
F
G
� 
5263
F l � 
H��
H \  � 
I
J
I l ��
K��
K n  ��
L
M
L 1  ���
� 
2903
M 1  ���
� 
sele�  �  
J m  ���� �  �  
G �
N�

� 
5264
N [  
O
P
O l 
Q�	�
Q n  
R
S
R 1  �
� 
2905
S 1  �
� 
sele�	  �  
P m  �� �
  �  �  �  
A 
T
U
T I ,�
V
W
� .sysodlogaskr        TEXT
V b  
X
Y
X m  
Z
Z �
[
[ X N o   B i b D e s k   r e f e r e n c e   m a t c h e s   t h i s   c i t e   k e y :  
Y o  ��  0 currentcitekey currentCiteKey
W �
\
]
� 
btns
\ J   
^
^ 
_�
_ m  
`
` �
a
a  O K�  
] � 
b��
�  
dflt
b J  #(
c
c 
d��
d m  #&
e
e �
f
f  O K��  ��  
U 
g��
g L  -/����  ��  
3 m  ��
h
h�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  
0   no match found   
1 �
i
i    n o   m a t c h   f o u n d�(  �'  
( 
j
k
j Z  5^
l
m��
n
l F  5N
o
p
o > 5;
q
r
q o  58���� "0 ibidcitekeylist ibidCiteKeyList
r J  8:����  
p = >J
s
t
s n  >F
u
v
u 4 AF��
w
�� 
cobj
w m  DE������
v o  >A���� "0 ibidcitekeylist ibidCiteKeyList
t o  FI����  0 currentcitekey currentCiteKey
m l QV
x
y
z
x r  QV
{
|
{ m  QR��
�� boovtrue
| o      ���� 0 ibidreference ibidReference
y !  check for ibid. references   
z �
}
} 6   c h e c k   f o r   i b i d .   r e f e r e n c e s��  
n r  Y^
~

~ m  YZ��
�� boovfals
 o      ���� 0 ibidreference ibidReference
k 
�
�
� r  _j
�
�
� b  _f
�
�
� o  _b���� "0 ibidcitekeylist ibidCiteKeyList
� o  be����  0 currentcitekey currentCiteKey
� o      ���� "0 ibidcitekeylist ibidCiteKeyList
� 
�
�
� Z  k�
�
�����
� H  ks
�
� E  kr
�
�
� o  kn���� 0 citekeylist citeKeyList
� o  nq����  0 currentcitekey currentCiteKey
� l v�
�
�
�
� r  v�
�
�
� b  v}
�
�
� o  vy���� 0 citekeylist citeKeyList
� o  y|����  0 currentcitekey currentCiteKey
� o      ���� 0 citekeylist citeKeyList
�   list all cite keys   
� �
�
� &   l i s t   a l l   c i t e   k e y s��  ��  
� 
�
�
� Y  ��
���
�
���
� l ��
�
�
�
� Z  ��
�
�����
� = ��
�
�
� n  ��
�
�
� 4  ����
�
�� 
cobj
� o  ������ 0 eachinteger eachInteger
� o  ������ 0 citekeylist citeKeyList
� o  ������  0 currentcitekey currentCiteKey
� r  ��
�
�
� o  ������ 0 eachinteger eachInteger
� o      ���� 0 citekeynumber citeKeyNumber��  ��  
� ) # get number for numbered references   
� �
�
� F   g e t   n u m b e r   f o r   n u m b e r e d   r e f e r e n c e s�� 0 eachinteger eachInteger
� m  ������ 
� l ��
�����
� I ����
���
�� .corecnte****       ****
� o  ������ 0 citekeylist citeKeyList��  ��  ��  ��  
� 
�
�
� l ��
�
�
�
� r  ��
�
�
� b  ��
�
�
� o  ������ ,0 multiplepublications multiplePublications
� o  ������ (0 currentpublication currentPublication
� o      ���� ,0 multiplepublications multiplePublications
� 2 , list multiple publications in current field   
� �
�
� X   l i s t   m u l t i p l e   p u b l i c a t i o n s   i n   c u r r e n t   f i e l d
� 
�
�
� l ����
�
���  
�   list of all publications   
� �
�
� 2   l i s t   o f   a l l   p u b l i c a t i o n s
� 
�
�
� Z ��
�
�����
� H  ��
�
� E  ��
�
�
� o  ������ "0 publicationlist publicationList
� o  ������ (0 currentpublication currentPublication
� r  ��
�
�
� b  ��
�
�
� o  ������ "0 publicationlist publicationList
� o  ������ (0 currentpublication currentPublication
� o      ���� "0 publicationlist publicationList��  ��  
� 
�
�
� l ����
�
���  
� 6 0 list references to be made italic such as films   
� �
�
� `   l i s t   r e f e r e n c e s   t o   b e   m a d e   i t a l i c   s u c h   a s   f i l m s
� 
�
�
� Z  �7
�
�����
� E  ��
�
�
� o  ������ 00 italicpublicationtypes italicPublicationTypes
� n  ��
�
�
� m  ����
�� 
type
� l ��
�����
� 6 ��
�
�
� n ��
�
�
� 2 ����
�� 
bibi
� o  ������ "0 bibdeskdocument bibDeskDocument
� = ��
�
�
� 1  ����
�� 
ckey
� o  ������ 0 citekey citeKey��  ��  
� k  �3
�
� 
�
�
� r  �
�
�
� n  �
�
�
� 1  ��
�� 
titl
� l �
�����
� 6 �
�
�
� n �
�
�
� 2 ���
�� 
bibi
� o  ������ "0 bibdeskdocument bibDeskDocument
� = 
�
�
� 1  	��
�� 
ckey
� o  
���� 0 citekey citeKey��  ��  
� o      ���� ,0 italicreferencetitle italicReferenceTitle
� 
���
� Z 3
�
�����
� H  !
�
� E   
�
�
� o  ���� $0 italicreferences italicReferences
� o  ���� ,0 italicreferencetitle italicReferenceTitle
� r  $/
�
�
� b  $+
�
�
� o  $'���� $0 italicreferences italicReferences
� o  '*���� ,0 italicreferencetitle italicReferenceTitle
� o      ���� $0 italicreferences italicReferences��  ��  ��  ��  ��  
� 
�
�
� r  8C
�
�
� b  8?
�
�
� o  8;���� $0 multiplecitekeys multipleCiteKeys
� o  ;>���� 0 citekeynumber citeKeyNumber
� o      ���� $0 multiplecitekeys multipleCiteKeys
� 
���
� r  DK
�
�
� o  DG���� $0 multiplecitekeys multipleCiteKeys
� o      ���� &0 citekeynumbertext citeKeyNumberText��  
 = 7 loop through every cite key and match to a publication   
 �
�
� n   l o o p   t h r o u g h   e v e r y   c i t e   k e y   a n d   m a t c h   t o   a   p u b l i c a t i o n�6 0 eachinteger eachInteger
 m  ������ 
 I ����
���
�� .corecnte****       ****
� n  ��
�
�
� 2 ����
�� 
citm
� o  ������ 0 citekey citeKey��  �5  �7  	� &   match cite keys to publications   	� �
�
� @   m a t c h   c i t e   k e y s   t o   p u b l i c a t i o n s	� m  wx
�
��                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  	� 
�
�
� Z  R� ��  = RY l RW���� c  RW o  RS���� (0 numberedreferences numberedReferences m  SV��
�� 
bool��  ��   m  WX��
�� boovfals k  \� 	
	 l \\����   $  format author-year references    � <   f o r m a t   a u t h o r - y e a r   r e f e r e n c e s
  Z  \��� E  \c o  \_���� 0 	fieldcode 	fieldCode m  _b � 
 c i t e p k  f�  Z  f��� = fk o  fi���� 0 ibidreference ibidReference m  ij��
�� boovtrue r  nu  m  nq!! �"" 
 i b i d .  o      ���� 0 templatedtext templatedText��   O  x�#$# r  |�%&% l |�'����' I |���()
�� .BDSKttxtnull���     docu( 4  |���*
�� 
docu* o  ~���� *0 bibdeskdocumentname bibDeskDocumentName) ��+,
�� 
usTx+ o  ������ 0 citeptemplate citepTemplate, ��-��
�� 
for - o  ������ ,0 multiplepublications multiplePublications��  ��  ��  & o      ���� 0 templatedtext templatedText$ m  xy..�                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��   /��/ r  ��010 b  ��232 b  ��454 b  ��676 b  ��898 o  ������ "0 parenthesisopen parenthesisOpen9 o  ������ "0 referenceprefix referencePrefix7 o  ������ 0 templatedtext templatedText5 o  ������ "0 referencesuffix referenceSuffix3 o  ������ $0 parenthesisclose parenthesisClose1 o      ���� (0 formattedreference formattedReference��   :;: E  ��<=< o  ������ 0 	fieldcode 	fieldCode= m  ��>> �?? 
 c i t e t; @A@ k  ��BB CDC Z  ��EF��GE = ��HIH o  ������ 0 ibidreference ibidReferenceI m  ����
�� boovtrueF r  ��JKJ m  ��LL �MM 
 i b i d .K o      ���� 0 templatedtext templatedText��  G O  ��NON r  ��PQP l ��R����R I ���ST
� .BDSKttxtnull���     docuS 4  ���~U
�~ 
docuU o  ���}�} *0 bibdeskdocumentname bibDeskDocumentNameT �|VW
�| 
usTxV o  ���{�{ 0 citettemplate citetTemplateW �zX�y
�z 
for X o  ���x�x ,0 multiplepublications multiplePublications�y  ��  ��  Q o      �w�w 0 templatedtext templatedTextO m  ��YY�                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  D Z�vZ r  ��[\[ b  ��]^] b  ��_`_ b  ��aba b  ��cdc o  ���u�u "0 parenthesisopen parenthesisOpend o  ���t�t "0 referenceprefix referencePrefixb o  ���s�s 0 templatedtext templatedText` o  ���r�r "0 referencesuffix referenceSuffix^ o  ���q�q $0 parenthesisclose parenthesisClose\ o      �p�p (0 formattedreference formattedReference�v  A efe E  �ghg o  ���o�o 0 	fieldcode 	fieldCodeh m  �ii �jj  c i t e a l pf klk k  Emm non Z  5pq�nrp = sts o  	�m�m 0 ibidreference ibidReferencet m  	
�l
�l boovtrueq r  uvu m  ww �xx 
 i b i d .v o      �k�k 0 templatedtext templatedText�n  r O  5yzy r  4{|{ l 0}�j�i} I 0�h~
�h .BDSKttxtnull���     docu~ 4   �g�
�g 
docu� o  �f�f *0 bibdeskdocumentname bibDeskDocumentName �e��
�e 
usTx� o  #&�d�d 0 citeptemplate citepTemplate� �c��b
�c 
for � o  ),�a�a ,0 multiplepublications multiplePublications�b  �j  �i  | o      �`�` 0 templatedtext templatedTextz m  ���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  o ��_� r  6E��� b  6A��� b  6=��� o  69�^�^ "0 referenceprefix referencePrefix� o  9<�]�] 0 templatedtext templatedText� o  =@�\�\ "0 referencesuffix referenceSuffix� o      �[�[ (0 formattedreference formattedReference�_  l ��� E  HO��� o  HK�Z�Z 0 	fieldcode 	fieldCode� m  KN�� ���  c i t e a l t� ��Y� k  R��� ��� Z  R����X�� = RW��� o  RU�W�W 0 ibidreference ibidReference� m  UV�V
�V boovtrue� r  Za��� m  Z]�� ��� 
 i b i d .� o      �U�U 0 templatedtext templatedText�X  � O  d���� r  h���� l h|��T�S� I h|�R��
�R .BDSKttxtnull���     docu� 4  hl�Q�
�Q 
docu� o  jk�P�P *0 bibdeskdocumentname bibDeskDocumentName� �O��
�O 
usTx� o  or�N�N 0 citettemplate citetTemplate� �M��L
�M 
for � o  ux�K�K ,0 multiplepublications multiplePublications�L  �T  �S  � o      �J�J 0 templatedtext templatedText� m  de���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  � ��I� r  ����� b  ����� b  ����� o  ���H�H "0 referenceprefix referencePrefix� o  ���G�G 0 templatedtext templatedText� o  ���F�F "0 referencesuffix referenceSuffix� o      �E�E (0 formattedreference formattedReference�I  �Y  ��   ��D� r  ����� o  ���C�C (0 formattedreference formattedReference� n      ��� 1  ���B
�B 
1650� n  ����� 1  ���A
�A 
2475� o  ���@�@ 0 currentfield currentField�D   ��� = ����� l ����?�>� c  ����� o  ���=�= (0 numberedreferences numberedReferences� m  ���<
�< 
bool�?  �>  � m  ���;
�; boovtrue� ��:� k  ���� ��� l ���9���9  � !  format numbered references   � ��� 6   f o r m a t   n u m b e r e d   r e f e r e n c e s� ��� r  ����� m  ���� ���  � o      �8�8 (0 formattedreference formattedReference� ��� r  ����� o  ���7�7 "0 numberseperator numberSeperator� n     ��� 1  ���6
�6 
txdl� 1  ���5
�5 
ascr� ��4� r  ����� b  ����� b  ����� o  ���3�3 "0 parenthesisopen parenthesisOpen� o  ���2�2 &0 citekeynumbertext citeKeyNumberText� o  ���1�1 $0 parenthesisclose parenthesisClose� n      ��� 1  ���0
�0 
1650� n  ����� 1  ���/
�/ 
2475� o  ���.�. 0 currentfield currentField�4  �:  ��  
� ��� l ���-���-  � ( " set superscript for current field   � ��� D   s e t   s u p e r s c r i p t   f o r   c u r r e n t   f i e l d� ��� Z  �����,� = ����� o  ���+�+ .0 superscriptreferences superscriptReferences� m  ���*
�* boovfals� r  ����� m  ���)
�) boovfals� n      ��� 1  ���(
�( 
2115� n  ����� 1  ���'
�' 
wFnO� n  ����� 1  ���&
�& 
2475� o  ���%�% 0 currentfield currentField� ��� = ����� o  ���$�$ .0 superscriptreferences superscriptReferences� m  ���#
�# boovtrue� ��"� r  ���� m  ���!
�! boovtrue� n      ��� 1  
� 
�  
2115� n  ���� 1  �
� 
wFnO� n  ���� 1  ��
� 
2475� o  ���� 0 currentfield currentField�"  �,  � � � r   m  �
� boovfals n       1  �
� 
2482 o  �� 0 currentfield currentField   l ��   $  set italics for current field    �		 <   s e t   i t a l i c s   f o r   c u r r e n t   f i e l d 

 r  # b   o  �� 0 italiclatin italicLatin o  �� $0 italicreferences italicReferences o      ��  0 italiccombined italicCombined � X  $�� Z  :��� E  :A o  :=�� (0 formattedreference formattedReference o  =@�� 0 
eachitalic 
eachItalic k  D�  r  DK m  DG �  , n      1  HJ�
� 
txdl 1  GH�
� 
ascr  !  X  L�"�#" k  b�$$ %&% r  bm'(' n  bi)*) 1  ei�
� 
1717* 1  be�
� 
sele( o      �
�
 0 
findobject 
findObject& +,+ r  ny-.- o  nq�	�	 0 
italicfind 
italicFind. n      /0/ 1  tx�
� 
16500 o  qt�� 0 
findobject 
findObject, 121 r  z�343 o  z}�� 0 
italicfind 
italicFind4 n      565 1  ���
� 
16506 n  }�787 m  ���
� 
w1258 o  }��� 0 
findobject 
findObject2 9:9 r  ��;<; m  ���
� boovtrue< n      =>= 1  ���
� 
ital> n  ��?@? 1  ��� 
�  
wFnO@ n  ��ABA m  ����
�� 
w125B o  ������ 0 
findobject 
findObject: C��C I ����DE
�� .sWRD1874null���     w124D o  ������ 0 
findobject 
findObjectE ��F��
�� 
5642F m  ����
�� e273� ��  ��  � 0 
italicfind 
italicFind# o  OR����  0 italiccombined italicCombined! G��G r  ��HIH m  ��JJ �KK  I n     LML 1  ����
�� 
txdlM 1  ����
�� 
ascr��  �  �  � 0 
eachitalic 
eachItalic o  '*����  0 italiccombined italicCombined�  	 ) # extract cite key and prefix/suffix   	� �NN F   e x t r a c t   c i t e   k e y   a n d   p r e f i x / s u f f i x�r  �q  �s  	N 3 - convert ADDIN fields to formatted references   	O �OO Z   c o n v e r t   A D D I N   f i e l d s   t o   f o r m a t t e d   r e f e r e n c e s�} 0 fieldnumber fieldNumber	K m  ������ 	L I ����P��
�� .corecnte****       ****P n  ��QRQ 2 ����
�� 
w170R o  ������ 0 formatrange formatRange��  �|  	I STS l ����UV��  U   format bibliography   V �WW (   f o r m a t   b i b l i o g r a p h yT XYX r  ��Z[Z I ����\��
�� .corecnte****       ****\ n  ��]^] 2 ����
�� 
w170^ o  ������ 0 formatrange formatRange��  [ o      ���� 0 
fieldcount 
fieldCountY _��_ W  �	�`a` k  �	�bb cdc r  ��efe n  ��ghg 4  ����i
�� 
w170i o  ������ 0 
fieldcount 
fieldCounth o  ������ 0 formatrange formatRangef o      ���� 0 currentfield currentFieldd jkj r  �lml n  ��non 1  ����
�� 
1650o n  ��pqp 1  ����
�� 
2469q o  ������ 0 currentfield currentFieldm o      ���� 0 	fieldcode 	fieldCodek rsr Z  	�tu����t F   vwv = xyx n  	z{z 1  	��
�� 
wFtP{ o  ���� 0 currentfield currentFieldy m  	��
�� e183G Qw E  |}| n  ~~ 4  ���
�� 
cwor� m  ����  o  ���� 0 	fieldcode 	fieldCode} m  �� ���  b i b l i o g r a p h yu k  #	��� ��� r  #*��� o  #&���� 0 currentfield currentField� o      ���� &0 bibliographyfield bibliographyField� ��� I +2�����
�� .miscslctnull���    obj � o  +.���� &0 bibliographyfield bibliographyField��  � ��� l 33������  � > 8 compile the bibliography text and export to a temp file   � ��� p   c o m p i l e   t h e   b i b l i o g r a p h y   t e x t   a n d   e x p o r t   t o   a   t e m p   f i l e� ��� O  3{��� k  9z�� ��� r  9L��� I 9H����
�� .earsffdralis        afdr� m  9<��
�� afdrtemp� �����
�� 
from� 1  ?D��
�� 
fldu��  � o      ���� 0 tempdirectory tempDirectory� ��� r  M[��� c  MW��� 4  MS���
�� 
psxf� o  QR���� 40 bibliographytemplatepath bibliographyTemplatePath� m  SV��
�� 
TEXT� o      ���� 40 bibliographytemplatefile bibliographyTemplateFile� ��� r  \j��� n  \f��� 1  bf��
�� 
pnam� 4  \b���
�� 
file� o  `a���� 40 bibliographytemplatepath bibliographyTemplatePath� o      ���� <0 bibliographytemplatefilename bibliographyTemplateFileName� ���� r  kz��� b  kv��� l kr������ c  kr��� o  kn���� 0 tempdirectory tempDirectory� m  nq��
�� 
TEXT��  ��  � o  ru���� <0 bibliographytemplatefilename bibliographyTemplateFileName� o      ���� 0 tempfile  ��  � m  36���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � ��� O  |���� k  ���� ��� Z  ��������� F  ����� > ����� o  ������ (0 bibliographysortby bibliographySortBy� m  ���� ���  � = ����� l �������� c  ����� o  ������ (0 numberedreferences numberedReferences� m  ����
�� 
bool��  ��  � m  ����
�� boovfals� k  ���� ��� I ������
�� .BDSKSortnull���     ****� o  ������ "0 publicationlist publicationList� �����
�� 
by  � o  ������ (0 bibliographysortby bibliographySortBy��  � ���� r  ����� 1  ����
�� 
rslt� o      ���� "0 publicationlist publicationList��  ��  ��  � ���� I ������
�� .BDSKexptnull���     ****� o  ������ "0 bibdeskdocument bibDeskDocument� ����
�� 
uset� 4  �����
�� 
file� l �������� c  ����� o  ������ 40 bibliographytemplatefile bibliographyTemplateFile� m  ����
�� 
TEXT��  ��  � ����
�� 
to  � 4  �����
�� 
file� o  ������ 0 tempfile  � �����
�� 
for � o  ������ "0 publicationlist publicationList��  ��  � m  |}���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  � ��� l ��������  � 7 1 insert exported text into the bibliography field   � ��� b   i n s e r t   e x p o r t e d   t e x t   i n t o   t h e   b i b l i o g r a p h y   f i e l d� ��� r  ����� m  ����
�� boovfals� n      ��� 1  ����
�� 
2482� o  ������ &0 bibliographyfield bibliographyField� ��� r  ����� m  ���� ���  *� n      ��� 1  ����
�� 
1650� n  ����� 1  ����
�� 
2475� o  ������ &0 bibliographyfield bibliographyField� ��� I �	�����
�� .sWRDwiFlnull��� ��� null��  � ����
�� 
insh� l �	������ I �	����
�� .sWRD1430null���     docu� o  ������ 0 worddocument wordDocument� ����
�� 
5263� l �	������ n  �	��� 1  �	��
�� 
wSCx� n  ��   1  ����
�� 
2475 o  ������ &0 bibliographyfield bibliographyField��  ��  � ����
�� 
5264 l 		���� n  		 1  	
	��
�� 
wSCx n  		
 1  		
��
�� 
2475 o  		���� &0 bibliographyfield bibliographyField��  ��  ��  ��  ��  � ����
�� 
5015 o  		�� 0 tempfile  ��  � 	
	 O 		0 I 	#	/�~�}
�~ .coredelonull���     ditm 4  	#	+�|
�| 
file o  	'	*�{�{ 0 tempfile  �}   m  		 �                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  
  l 	1	1�z�z   ; 5 set Word style of bibliography and force font object    � j   s e t   W o r d   s t y l e   o f   b i b l i o g r a p h y   a n d   f o r c e   f o n t   o b j e c t  r  	1	> o  	1	2�y�y &0 bibliographystyle bibliographyStyle n       1  	9	=�x
�x 
1695 n  	2	9 1  	5	9�w
�w 
2475 o  	2	5�v�v &0 bibliographyfield bibliographyField  r  	?	]  n  	?	M!"! 1  	I	M�u
�u 
pnam" n  	?	I#$# 1  	E	I�t
�t 
wFnO$ n  	?	E%&% 4  	@	E�s'
�s 
w173' o  	C	D�r�r &0 bibliographystyle bibliographyStyle& o  	?	@�q�q 0 worddocument wordDocument  n      ()( 1  	X	\�p
�p 
pnam) n  	M	X*+* 1  	T	X�o
�o 
wFnO+ n  	M	T,-, 1  	P	T�n
�n 
2475- o  	M	P�m�m &0 bibliographyfield bibliographyField ./. l 	^	^�l01�l  0 $  remove superfluous characters   1 �22 <   r e m o v e   s u p e r f l u o u s   c h a r a c t e r s/ 343 I 	^	��k5�j
�k .coredelonull���     ditm5 l 	^	}6�i�h6 I 	^	}�g78
�g .sWRD1430null���     docu7 o  	^	_�f�f 0 worddocument wordDocument8 �e9:
�e 
52639 l 	`	m;�d�c; \  	`	m<=< l 	`	k>�b�a> n  	`	k?@? 1  	g	k�`
�` 
1656@ n  	`	gABA 1  	c	g�_
�_ 
2475B o  	`	c�^�^ &0 bibliographyfield bibliographyField�b  �a  = m  	k	l�]�] �d  �c  : �\C�[
�\ 
5264C l 	n	yD�Z�YD n  	n	yEFE 1  	u	y�X
�X 
1656F n  	n	uGHG 1  	q	u�W
�W 
2475H o  	n	q�V�V &0 bibliographyfield bibliographyField�Z  �Y  �[  �i  �h  �j  4 I�UI  S  	�	��U  ��  ��  s J�TJ r  	�	�KLK \  	�	�MNM o  	�	��S�S 0 
fieldcount 
fieldCountN m  	�	��R�R L o      �Q�Q 0 
fieldcount 
fieldCount�T  a = ��OPO o  ���P�P 0 
fieldcount 
fieldCountP m  ���O�O  ��  \ m  QQ�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  Z RSR l 	�	��NTU�N  T , & restore Word to its original position   U �VV L   r e s t o r e   W o r d   t o   i t s   o r i g i n a l   p o s i t i o nS WXW O  	�	�YZY k  	�	�[[ \]\ Q  	�	�^_`^ I 	�	��Ma�L
�M .miscslctnull���    obj a o  	�	��K�K  0 selectionrange selectionRange�L  _ R      �J�I�H
�J .ascrerr ****      � ****�I  �H  ` l 	�	�bcdb I 	�	��Ge�F
�G .miscslctnull���    obj e l 	�	�f�E�Df I 	�	��Cgh
�C .sWRD1430null���     docug o  	�	��B�B 0 worddocument wordDocumenth �Aij
�A 
5263i l 	�	�k�@�?k n  	�	�lml 1  	�	��>
�> 
2905m 1  	�	��=
�= 
sele�@  �?  j �<n�;
�< 
5264n l 	�	�o�:�9o n  	�	�pqp 1  	�	��8
�8 
2905q 1  	�	��7
�7 
sele�:  �9  �;  �E  �D  �F  c > 8 produces error if cursor was at the end of the document   d �rr p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n t] s�6s r  	�	�tut m  	�	�vv �ww < F i n i s h e d   f o r m a t t i n g   r e f e r e n c e su 1  	�	��5
�5 
1091�6  Z m  	�	�xx�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  X y�4y L  	�	�zz o  	�	��3�3 ,0 bibliographysettings bibliographySettings�4  � {|{ l     �2�1�0�2  �1  �0  | }~} l     �/��/    -   � ���  -~ ��� l     �.�-�,�.  �-  �,  � ��� i    ��� I      �+�*�)�+ (0 unformatreferences unformatReferences�*  �)  � k    ��� ��� l     �(���(  �   global setup   � ���    g l o b a l   s e t u p� ��� O     $��� k    #�� ��� r    	��� 1    �'
�' 
1003� o      �&�& 0 worddocument wordDocument� ��� r   
 ��� I  
 �%��
�% .sWRD1430null���     docu� o   
 �$�$ 0 worddocument wordDocument� �#��
�# 
5263� l   ��"�!� n    ��� 1    � 
�  
2903� 1    �
� 
sele�"  �!  � ���
� 
5264� l   ���� n    ��� 1    �
� 
2905� 1    �
� 
sele�  �  �  � o      ��  0 selectionrange selectionRange� ��� r    #��� m    �� ��� 4 U n f o r m a t t i n g   r e f e r e n c e s . . .� 1    "�
� 
1091�  � m     ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� r   % *��� m   % &�� ���  � n     ��� 1   ' )�
� 
txdl� 1   & '�
� 
ascr� ��� l  + +����  �   unformat references   � ��� (   u n f o r m a t   r e f e r e n c e s� ��� O   +R��� k   /Q�� ��� I  / 4���
� .miscactvnull��� ��� null�  �  � ��� Z   5 p����� =  5 @��� l  5 :���� n   5 :��� 1   8 :�
� 
2903� 1   5 8�
� 
sele�  �  � l  : ?��
�	� n   : ?��� 1   = ?�
� 
2905� 1   : =�
� 
sele�
  �	  � k   C j�� ��� r   C F��� o   C D�� 0 worddocument wordDocument� o      �� 0 formatrange formatRange� ��� Q   G j���� I  J `���
� .sysodlogaskr        TEXT� m   J K�� ��� j A r e   y o u   s u r e ?   T h i s   w i l l   u n f o r m a t   t h e   e n t i r e   d o c u m e n t !� ���
� 
btns� J   N V�� ��� m   N Q�� ���  C a n c e l� ��� m   Q T�� ���  O K�  � � ���
�  
dflt� m   Y \�� ���  O K��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   h j����  �  �  � r   m p��� o   m n����  0 selectionrange selectionRange� o      ���� 0 formatrange formatRange� ��� r   q x��� n   q v��� 2  r v��
�� 
w170� o   q r���� 0 formatrange formatRange� o      ���� 0 	fieldlist 	fieldList� ��� Z   y �������� =  y ���� l  y ~������ I  y ~�����
�� .corecnte****       ****� o   y z���� 0 	fieldlist 	fieldList��  ��  ��  � m   ~ ����  � k   � ��� ��� I  � �����
�� .sysodlogaskr        TEXT� m   � ��� ��� f T h e r e   a r e   n o   f o r m a t t e d   r e f e r e n c e s   i n   t h i s   s e l e c t i o n� ����
�� 
btns� J   � �   �� m   � � �  O K��  � ����
�� 
dflt m   � � �  O K��  � �� L   � �����  ��  ��  ��  � �� X   �Q	��
	 k   �L  r   � � n   � � 1   � ���
�� 
1650 n   � � 1   � ���
�� 
2469 o   � ����� 0 currentfield currentField o      ���� 0 	fieldcode 	fieldCode �� Z   �L���� F   � � =  � � n   � � 1   � ���
�� 
wFtP o   � ����� 0 currentfield currentField m   � ���
�� e183G Q E   � � J   � �  !  m   � �"" �##  c i t e! $%$ m   � �&& �'' 
 c i t e p% ()( m   � �** �++ 
 c i t e t) ,-, m   � �.. �//  c i t e a l p- 010 m   � �22 �33  c i t e a l t1 4��4 m   � �55 �66  b i b l i o g r a p h y��   n   � �787 4   � ���9
�� 
cwor9 m   � ����� 8 o   � ����� 0 	fieldcode 	fieldCode k   �H:: ;<; Z  �=>����= E   � �?@? o   � ����� 0 	fieldcode 	fieldCode@ m   � �AA �BB  b i b l i o g r a p h y> r   �	CDC m   � �EE �FF  N o r m a lD n      GHG 1  ��
�� 
1695H n   �IJI 1   ��
�� 
2475J o   � ���� 0 currentfield currentField��  ��  < KLK r  #MNM c  !OPO n  QRQ 7 ��ST
�� 
cha S m  ���� T m  ������R o  ���� 0 	fieldcode 	fieldCodeP m   ��
�� 
TEXTN o      ���� 0 	fieldcode 	fieldCodeL UVU I $)��W��
�� .miscslctnull���    obj W o  $%���� 0 currentfield currentField��  V XYX I *B����Z
�� .sWRDwITRnull��� ��� null��  Z ��[\
�� 
5418[ b  .3]^] m  .1__ �``  \^ o  12���� 0 	fieldcode 	fieldCode\ ��a��
�� 
insha n  6>bcb  ;  =>c n  6=ded 1  9=��
�� 
wTxRe 1  69��
�� 
sele��  Y f��f I CH��g��
�� .coredelonull���     ditmg o  CD���� 0 currentfield currentField��  ��  ��  ��  ��  �� 0 currentfield currentField
 n   � �hih 1   � ���
�� 
rvsei o   � ����� 0 	fieldlist 	fieldList��  � m   + ,jj�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � klk l SS��mn��  m , & restore Word to its original position   n �oo L   r e s t o r e   W o r d   t o   i t s   o r i g i n a l   p o s i t i o nl p��p O  S�qrq k  W�ss tut Q  W|vwxv I Z_��y��
�� .miscslctnull���    obj y o  Z[����  0 selectionrange selectionRange��  w R      ������
�� .ascrerr ****      � ****��  ��  x l g|z{|z I g|��}��
�� .miscslctnull���    obj } l gx~����~ I gx���
�� .sWRD1430null���     docu o  gh���� 0 worddocument wordDocument� ����
�� 
5263� l in������ n  in��� 1  ln��
�� 
2905� 1  il��
�� 
sele��  ��  � �����
�� 
5264� l ot������ n  ot��� 1  rt��
�� 
2905� 1  or��
�� 
sele��  ��  ��  ��  ��  ��  { > 8 produces error if cursor was at the end of the document   | ��� p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n tu ���� r  }���� m  }��� ��� @ F i n i s h e d   u n f o r m a t t i n g   r e f e r e n c e s� 1  ����
�� 
1091��  r m  ST���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  �  -   � ���  -� ��� l     ��������  ��  ��  � ��� i    ��� I      �������� *0 fillemptyreferences fillEmptyReferences��  ��  � k    ��� ��� l     ������  � , & find empty references contained by {}   � ��� L   f i n d   e m p t y   r e f e r e n c e s   c o n t a i n e d   b y   { }� ��� r     ��� m     �� ���  � n     ��� 1    ��
�� 
txdl� 1    ��
�� 
ascr� ��� O   D��� k   
C�� ��� r   
 ��� 1   
 ��
�� 
1003� o      ���� 0 worddocument wordDocument� ��� r    #��� l   !������ I   !����
�� .sWRD1430null���     docu� o    ���� 0 worddocument wordDocument� ����
�� 
5263� l   ������ n    ��� 1    ��
�� 
2903� 1    ��
�� 
sele��  ��  � �����
�� 
5264� l   ������ n    ��� 1    ��
�� 
2905� 1    ��
�� 
sele��  ��  ��  ��  ��  � o      ����  0 selectionrange selectionRange� ��� r   $ )��� m   $ %�� ��� 6 F i l l i n g   e m p t y   r e f e r e n c e s . . .� 1   % (�
� 
1091� ��� l  * *�~���~  � / ) start from the beginning of the document   � ��� R   s t a r t   f r o m   t h e   b e g i n n i n g   o f   t h e   d o c u m e n t� ��� I  * 7�}��|
�} .miscslctnull���    obj � l  * 3��{�z� I  * 3�y��
�y .sWRD1430null���     docu� o   * +�x�x 0 worddocument wordDocument� �w��
�w 
5263� m   , -�v�v  � �u��t
�u 
5264� m   . /�s�s  �t  �{  �z  �|  � ��� l  8 8�r���r  � %  set up the find system in Word   � ��� >   s e t   u p   t h e   f i n d   s y s t e m   i n   W o r d� ��� r   8 ?��� n   8 =��� 1   ; =�q
�q 
1717� 1   8 ;�p
�p 
sele� o      �o�o 0 
findobject 
findObject� ��� r   @ G��� m   @ A�n
�n boovtrue� n      ��� 1   B F�m
�m 
1840� o   A B�l�l 0 
findobject 
findObject� ��� l  H H�k���k  � - ' prevent wrapping to enable skip option   � ��� N   p r e v e n t   w r a p p i n g   t o   e n a b l e   s k i p   o p t i o n� ��� r   H Q��� m   H K�j
�j e265�  � n      ��� 1   L P�i
�i 
1862� o   K L�h�h 0 
findobject 
findObject� ��� l  R R�g���g  � ? 9 find empty references and contents but not cite commands   � ��� r   f i n d   e m p t y   r e f e r e n c e s   a n d   c o n t e n t s   b u t   n o t   c i t e   c o m m a n d s� ��� r   R [��� m   R U�� ���  [ ! \ ] ] \ { * \ }� n         1   V Z�f
�f 
1650 o   U V�e�e 0 
findobject 
findObject�  r   \ _ m   \ ]�d�d   o      �c�c 0 
thecounter 
theCounter  V   `1	 k   j,

  r   j w b   j s b   j o m   j m � F S e a r c h i n g   f o r   e m p t y   r e f e r e n c e s . . .   ( o   m n�b�b 0 
thecounter 
theCounter m   o r �  ) 1   s v�a
�a 
1091  r   x � n   x � 1   � ��`
�` 
1650 l  x ��_�^ I  x ��]
�] .sWRD1430null���     docu o   x y�\�\ 0 worddocument wordDocument �[ !
�[ 
5263  [   z �"#" l  z $�Z�Y$ n   z %&% 1   } �X
�X 
2903& 1   z }�W
�W 
sele�Z  �Y  # m    ��V�V ! �U'�T
�U 
5264' l  � �(�S�R( n   � �)*) 1   � ��Q
�Q 
2905* 1   � ��P
�P 
sele�S  �R  �T  �_  �^   o      �O�O (0 emptyreferencetext emptyReferenceText +,+ r   � �-.- b   � �/0/ b   � �121 m   � �33 �44  "2 n   � �565 7 � ��N78
�N 
cha 7 m   � ��M�M 8 m   � ��L�L��6 l  � �9�K�J9 c   � �:;: n   � �<=< 1   � ��I
�I 
1650= n   � �>?> 2  � ��H
�H 
csen? 1   � ��G
�G 
sele; m   � ��F
�F 
TEXT�K  �J  0 m   � �@@ �AA  ". o      �E�E .0 emptyreferencecontext emptyReferenceContext, BCB Q   �DEFD k   �	GG HIH r   � �JKJ J   � �LL MNM m   � �OO �PP  ( A u t h o r ,   Y E A R )N QRQ m   � �SS �TT  ( Y E A R )R UVU m   � �WW �XX  ( C u s t o m . . . )V Y�DY m   � �ZZ �[[ $ S k i p   t h i s   i n s t a n c e�D  K J      \\ ]^] o      �C�C 0 	citeplist 	citepList^ _`_ o      �B�B 0 	citetlist 	citetList` aba o      �A�A 0 citealplist citealpListb c�@c o      �?�? 0 citeskip citeSkip�@  I ded I  ��>fg
�> .gtqpchltns    @   @ ns  f J   � �hh iji o   � ��=�= 0 	citeplist 	citepListj klk o   � ��<�< 0 	citetlist 	citetListl mnm o   � ��;�; 0 citealplist citealpListn o�:o o   � ��9�9 0 citeskip citeSkip�:  g �8pq
�8 
apprp m   � �rr �ss 0 F i l l   E m p t y   R e f e r e n c e s   { }q �7tu
�7 
prmpt b   	vwv b   xyx m   zz �{{ X P l e a s e   s e l e c t   r e f e r e n c e [ s ]   i n   B i b D e s k   f o r : 
 
y o  �6�6 .0 emptyreferencecontext emptyReferenceContextw m  || �}}  
u �5~�4
�5 
inSL~ o  �3�3 0 	citeplist 	citepList�4  e � r  ��� 1  �2
�2 
rslt� o      �1�1 0 	theresult 	theResult� ��0� Z  	���/�.� H  �� E  ��� o  �-�- 0 	theresult 	theResult� o  �,�, 0 citeskip citeSkip� k  !�� ��� Z  !�����+� E  !$��� o  !"�*�* 0 	theresult 	theResult� o  "#�)�) 0 	citeplist 	citepList� O 'K��� r  -J��� I -H�(��
�( .BDSKttxtnull���     docu� 4 -3�'�
�' 
docu� m  12�&�& � �%��
�% 
usTx� m  69�� ��� � < $ p u b l i c a t i o n s > < $ # = 1 ? > \ c i t e p [ ] [ ] { < $ c i t e K e y / > < ? $ # ? > , < $ c i t e K e y / > < / $ # ? > < / $ p u b l i c a t i o n s > }� �$��#
�$ 
for � n  <D��� 1  BD�"
�" 
sele� 4 <B�!�
�! 
docu� m  @A� �  �#  � o      �� 0 templatedtext templatedText� m  '*���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  � ��� E  NQ��� o  NO�� 0 	theresult 	theResult� o  OP�� 0 	citetlist 	citetList� ��� O Tx��� r  Zw��� I Zu���
� .BDSKttxtnull���     docu� 4 Z`��
� 
docu� m  ^_�� � ���
� 
usTx� m  cf�� ��� � < $ p u b l i c a t i o n s > < $ # = 1 ? > \ c i t e t [ ] [ ] { < $ c i t e K e y / > < ? $ # ? > , < $ c i t e K e y / > < / $ # ? > < / $ p u b l i c a t i o n s > }� ���
� 
for � n  iq��� 1  oq�
� 
sele� 4 io��
� 
docu� m  mn�� �  � o      �� 0 templatedtext templatedText� m  TW���                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  � ��� E  {~��� o  {|�� 0 	theresult 	theResult� o  |}�� 0 citealplist citealpList� ��� k  ���� ��� r  ����� b  ����� o  ���� .0 emptyreferencecontext emptyReferenceContext� m  ���� ��� v 
 	 
 ( U s e   ' c i t e a l p '   f o r   A u t h o r ,   Y E A R ,   a n d   ' c i t e a l t '   f o r   Y E A R )� o      �� .0 emptyreferencecontext emptyReferenceContext� ��� O ����� r  ����� I �����
� .BDSKttxtnull���     docu� 4 ����
� 
docu� m  ���� � �
��
�
 
usTx� m  ���� ��� � < $ p u b l i c a t i o n s > < $ # = 1 ? > \ c i t e a l p [ ] [ ] { < $ c i t e K e y / > } < ? $ # ? > \ c i t e a l p [ ;   ] [ ] { < $ c i t e K e y / > } < / $ # ? > < / $ p u b l i c a t i o n s >� �	��
�	 
for � n  ����� 1  ���
� 
sele� 4 ����
� 
docu� m  ���� �  � o      �� 0 templatedtext templatedText� m  �����                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  �  � ��� = ����� o  ���� 0 	theresult 	theResult� m  ���
� boovfals� ���  S  ���  �+  � ��� I ��� ��
�  .sysodlogaskr        TEXT� b  ����� m  ���� ��� L P l e a s e   c o n f i r m   t h e   c i t e   c o m m a n d   f o r : 
 
� o  ������ .0 emptyreferencecontext emptyReferenceContext� ����
�� 
appr� m  ���� ��� 0 F i l l   E m p t y   R e f e r e n c e s   { }� ����
�� 
dtxt� o  ������ 0 templatedtext templatedText� �����
�� 
dflt� m  ���� ���  O K��  � ��� r  ����� n  ����� 1  ����
�� 
ttxt� 1  ����
�� 
rslt� o      ���� 0 templatedtext templatedText� ��� r  ����� l �������� c  ����� o  ������ 0 templatedtext templatedText� m  ����
�� 
TEXT��  ��  � n      ��� 1  ����
�� 
1650� l �������� I ������
�� .sWRD1430null���     docu� o  ������ 0 worddocument wordDocument� ��� 
�� 
5263� [  �� l ������ n  �� 1  ����
�� 
2903 1  ����
�� 
sele��  ��   m  ������   ����
�� 
5264 l ������ n  ��	 1  ����
�� 
2905	 1  ����
�� 
sele��  ��  ��  ��  ��  � 
��
 r    [    o   ���� 0 
thecounter 
theCounter m  ����  o      ���� 0 
thecounter 
theCounter��  �/  �.  �0  E R      ������
�� .ascrerr ****      � ****��  ��  F k    l ����   + % prevent "no empty references" dialog    � J   p r e v e n t   " n o   e m p t y   r e f e r e n c e s "   d i a l o g  r   m  ������ o      ���� 0 
thecounter 
theCounter  l ����   N H exit repeat instead of return to avoid breaking the find system in Word    � �   e x i t   r e p e a t   i n s t e a d   o f   r e t u r n   t o   a v o i d   b r e a k i n g   t h e   f i n d   s y s t e m   i n   W o r d ��  S  ��  C   l ��!"��  ! ; 5 continue search from end of filled reference onwards   " �## j   c o n t i n u e   s e a r c h   f r o m   e n d   o f   f i l l e d   r e f e r e n c e   o n w a r d s  $��$ I ,��%��
�� .miscslctnull���    obj % l (&����& I (��'(
�� .sWRD1430null���     docu' o  ���� 0 worddocument wordDocument( ��)*
�� 
5263) l +����+ n  ,-, 1  ��
�� 
2905- 1  ��
�� 
sele��  ��  * ��.��
�� 
5264. l $/����/ n  $010 1  "$��
�� 
29051 1  "��
�� 
sele��  ��  ��  ��  ��  ��  ��  	 l  d i2����2 I  d i��3��
�� .sWRD1874null���     w1243 o   d e���� 0 
findobject 
findObject��  ��  ��   4��4 Q  2C56��5 I 5:��7��
�� .miscslctnull���    obj 7 o  56����  0 selectionrange selectionRange��  6 R      ������
�� .ascrerr ****      � ****��  ��  ��  ��  � m    88�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � 9:9 Z  E�;<=��; = EH>?> o  EF���� 0 
thecounter 
theCounter? m  FG����  < O  Ke@A@ I Od��BC
�� .sysodlogaskr        TEXTB m  ORDD �EE 4 N o   e m p t y   r e f e r e n c e s   f i l l e dC ��FG
�� 
btnsF J  UZHH I��I m  UXJJ �KK  O K��  G ��L��
�� 
dfltL m  ]`MM �NN  O K��  A m  KLOO�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  = PQP ? hkRSR o  hi���� 0 
thecounter 
theCounterS m  ij����  Q T��T Q  n�UV��U O  q�WXW k  u�YY Z[Z I u���\]
�� .sysodlogaskr        TEXT\ m  ux^^ �__ ^ E m p t y   r e f e r e n c e s   f i l l e d ,   f o r m a t   r e f e r e n c e s   n o w ?] ��`a
�� 
btns` J  {�bb cdc m  {~ee �ff  N od g��g m  ~�hh �ii  Y e s��  a ��jk
�� 
cbtnj m  ��ll �mm  N ok ��n��
�� 
dfltn m  ��oo �pp  Y e s��  [ q��q Z ��rs����r E  ��tut n  ��vwv 1  ����
�� 
bhitw 1  ����
�� 
rsltu m  ��xx �yy  Y e ss I  ���������� $0 menuchoiceformat menuChoiceFormat��  ��  ��  ��  ��  X m  qrzz�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  V R      ������
�� .ascrerr ****      � ****��  ��  ��  ��  ��  : {|{ O  ��}~} r  ��� m  ���� ��� B F i n i s h e d   f i l l i n g   e m p t y   r e f e r e n c e s� 1  ����
�� 
1091~ m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  | ���� L  ������  ��  � ��� l     ��������  ��  ��  � ��� l     ������  �  -   � ���  -� ��� l     ��������  ��  ��  � ��� i    ��� I      �������� 0 	toolsmenu 	toolsMenu��  ��  � k    ��� ��� l     ������  �   global setup   � ���    g l o b a l   s e t u p� ��� O      ��� k    �� ��� r    	��� 1    ��
�� 
1003� o      ���� 0 worddocument wordDocument� ���� r   
 ��� I  
 ����
�� .sWRD1430null���     docu� 1   
 ��
�� 
1003� ����
�� 
5263� l   ������ n    ��� 1    ��
�� 
2903� 1    ��
�� 
sele��  ��  � ����
�� 
5264� l   ��~�}� n    ��� 1    �|
�| 
2905� 1    �{
�{ 
sele�~  �}  �  � o      �z�z  0 selectionrange selectionRange��  � m     ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� l  ! !�y���y  �   open tools menu   � ���     o p e n   t o o l s   m e n u� ��� T   !��� k   &��� ��� O   & i��� k   * h�� ��� I  * /�x�w�v
�x .miscactvnull��� ��� null�w  �v  � ��� I  0 b�u��
�u .gtqpchltns    @   @ ns  � J   0 :�� ��� m   0 1�� ��� 8 C o n v e r t   f r o m   \ c i t e   t o   \ c i t e p� ��� m   1 2�� ��� 6 C o v e r t   f r o m   \ c i t e p   t o   \ c i t e� ��� m   2 3�� ��� & C o n v e r t   p . / p p .   t o   :� ��� m   3 4�� ��� & C o n v e r t   :   t o   p . / p p .� ��� m   4 5�� ��� " A d d   L e a d i n g   S p a c e� ��t� m   5 6�� ��� ( R e m o v e   L e a d i n g   S p a c e�t  � �s��
�s 
appr� m   = @�� ���   W o r d S c r i p t   T o o l s� �r��
�r 
prmp� m   C F�� ��� @ Y o u   s h o u l d   s a v e   y o u r   w o r k   f i r s t !� �q��
�q 
inSL� m   I L�� ��� 8 C o n v e r t   f r o m   \ c i t e   t o   \ c i t e p� �p��
�p 
cnbt� J   O T�� ��o� m   O R�� ���  E x i t�o  � �n��m
�n 
okbt� J   W \�� ��l� m   W Z�� ���  C h o o s e�l  �m  � ��k� r   c h��� 1   c f�j
�j 
rslt� o      �i�i 0 
menuchoice 
menuChoice�k  � m   & '���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� Z   j u���h�g� =  j m��� o   j k�f�f 0 
menuchoice 
menuChoice� m   k l�e
�e boovfals�  S   p q�h  �g  � ��� Z   v[� �d� E   v { o   v w�c�c 0 
menuchoice 
menuChoice m   w z � , E x p o r t   t o   S t a t i c   G r o u p  k   ~ ~  l  ~ ~�b�a�`�b  �a  �`   	
	 l  ~ ~�_�_     rewrite from scratch    � *   r e w r i t e   f r o m   s c r a t c h
 �^ l  ~ ~�]�\�[�]  �\  �[  �^    E   � � o   � ��Z�Z 0 
menuchoice 
menuChoice m   � � � 8 C o n v e r t   f r o m   \ c i t e   t o   \ c i t e p  k   �1  r   � � m   � � �   n      1   � ��Y
�Y 
txdl 1   � ��X
�X 
ascr  !  O   �."#" k   �-$$ %&% I  � ��W�V�U
�W .miscactvnull��� ��� null�V  �U  & '(' r   � �)*) n   � �+,+ 1   � ��T
�T 
1717, n   � �-.- 1   � ��S
�S 
wTxR. o   � ��R�R 0 worddocument wordDocument* o      �Q�Q 0 
findobject 
findObject( /0/ r   � �121 m   � ��P
�P boovfals2 n      343 1   � ��O
�O 
18404 o   � ��N�N 0 
findobject 
findObject0 5�M5 X   �-6�L76 k   �(88 9:9 r   � �;<; o   � ��K�K 0 citecommand citeCommand< n      =>= 1   � ��J
�J 
1650> o   � ��I�I 0 
findobject 
findObject: ?�H? Z   �(@AB�G@ =  � �CDC c   � �EFE o   � ��F�F 0 citecommand citeCommandF m   � ��E
�E 
TEXTD m   � �GG �HH  \ c i t e {A k   � �II JKJ r   � �LML m   � �NN �OO  \ c i t e p {M n      PQP 1   � ��D
�D 
1650Q n   � �RSR m   � ��C
�C 
w125S o   � ��B�B 0 
findobject 
findObjectK T�AT I  � ��@UV
�@ .sWRD1874null���     w124U o   � ��?�? 0 
findobject 
findObjectV �>W�=
�> 
5642W m   � ��<
�< e273� �=  �A  B XYX =  �Z[Z c   �\]\ o   � �;�; 0 citecommand citeCommand] m   �:
�: 
TEXT[ m  ^^ �__  \ c i t e [Y `�9` k  $aa bcb r  ded m  ff �gg  \ c i t e p [e n      hih 1  �8
�8 
1650i n  jkj m  �7
�7 
w125k o  �6�6 0 
findobject 
findObjectc l�5l I $�4mn
�4 .sWRD1874null���     w124m o  �3�3 0 
findobject 
findObjectn �2o�1
�2 
5642o m   �0
�0 e273� �1  �5  �9  �G  �H  �L 0 citecommand citeCommand7 J   � �pp qrq m   � �ss �tt  \ c i t e {r u�/u m   � �vv �ww  \ c i t e [�/  �M  # m   � �xx�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  ! y�.y L  /1�-�-  �.   z{z E  49|}| o  45�,�, 0 
menuchoice 
menuChoice} m  58~~ � 6 C o v e r t   f r o m   \ c i t e p   t o   \ c i t e{ ��� k  <��� ��� r  <G��� m  <?�� ���  � n     ��� 1  BF�+
�+ 
txdl� 1  ?B�*
�* 
ascr� ��� O  H���� k  L��� ��� I LQ�)�(�'
�) .miscactvnull��� ��� null�(  �'  � ��� r  R]��� n  R[��� 1  W[�&
�& 
1717� n  RW��� 1  SW�%
�% 
wTxR� o  RS�$�$ 0 worddocument wordDocument� o      �#�# 0 
findobject 
findObject� ��� r  ^e��� m  ^_�"
�" boovfals� n      ��� 1  `d�!
�! 
1840� o  _`� �  0 
findobject 
findObject� ��� X  f����� k  ���� ��� r  ����� o  ���� 0 citecommand citeCommand� n      ��� 1  ���
� 
1650� o  ���� 0 
findobject 
findObject� ��� Z  ������� = ����� c  ����� o  ���� 0 citecommand citeCommand� m  ���
� 
TEXT� m  ���� ���  \ c i t e p {� k  ���� ��� r  ����� m  ���� ���  \ c i t e {� n      ��� 1  ���
� 
1650� n  ����� m  ���
� 
w125� o  ���� 0 
findobject 
findObject� ��� I �����
� .sWRD1874null���     w124� o  ���� 0 
findobject 
findObject� ���
� 
5642� m  ���
� e273� �  �  � ��� = ����� c  ����� o  ���� 0 citecommand citeCommand� m  ���
� 
TEXT� m  ���� ���  \ c i t e p [� ��� k  ���� ��� r  ����� m  ���� ���  \ c i t e [� n      ��� 1  ���

�
 
1650� n  ����� m  ���	
�	 
w125� o  ���� 0 
findobject 
findObject� ��� I �����
� .sWRD1874null���     w124� o  ���� 0 
findobject 
findObject� ���
� 
5642� m  ���
� e273� �  �  �  �  �  � 0 citecommand citeCommand� J  iq�� ��� m  il�� ���  \ c i t e p {� ��� m  lo�� ���  \ c i t e p [�  �  � m  HI���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � �� � L  ������  �   � ��� E  ����� o  ������ 0 
menuchoice 
menuChoice� m  ���� ��� & C o n v e r t   p . / p p .   t o   :� ��� k  ���� ��� r  ����� m  ���� ���  � n     ��� 1  ����
�� 
txdl� 1  ����
�� 
ascr� ��� O  ����� k  ���� ��� I �������
�� .miscactvnull��� ��� null��  ��  � ��� r  � � n   1  	��
�� 
1717 n  	 1  	��
�� 
wTxR o  ���� 0 worddocument wordDocument  o      ���� 0 
findobject 
findObject�  r   m  ��
�� boovfals n      	
	 1  ��
�� 
1840
 o  ���� 0 
findobject 
findObject �� X  ��� k  3�  r  3: o  34���� 0 citecommand citeCommand n       1  59��
�� 
1650 o  45���� 0 
findobject 
findObject �� Z  ;��� = ;D c  ;@ o  ;<���� 0 citecommand citeCommand m  <?��
�� 
TEXT m  @C � 
 [ ,   p . k  G`  !  l GT"#$" r  GT%&% m  GJ'' �((  [ :  & n      )*) 1  OS��
�� 
1650* n  JO+,+ m  KO��
�� 
w125, o  JK���� 0 
findobject 
findObject#   normal colon   $ �--    n o r m a l   c o l o n! .��. I U`��/0
�� .sWRD1874null���     w124/ o  UV���� 0 
findobject 
findObject0 ��1��
�� 
56421 m  Y\��
�� e273� ��  ��   232 = cl454 c  ch676 o  cd���� 0 citecommand citeCommand7 m  dg��
�� 
TEXT5 m  hk88 �99  [ ,   p p .3 :��: k  o�;; <=< l o|>?@> r  o|ABA m  orCC �DD  [��  B n      EFE 1  w{��
�� 
1650F n  rwGHG m  sw��
�� 
w125H o  rs���� 0 
findobject 
findObject?   NBSp colon   @ �II    N B S p   c o l o n= J��J I }���KL
�� .sWRD1874null���     w124K o  }~���� 0 
findobject 
findObjectL ��M��
�� 
5642M m  ����
�� e273� ��  ��  ��  ��  ��  �� 0 citecommand citeCommand J  #NN OPO m  QQ �RR 
 [ ,   p .P S��S m  !TT �UU  [ ,   p p .��  ��  � m  ��VV�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � W��W L  ������  ��  � XYX E  ��Z[Z o  ������ 0 
menuchoice 
menuChoice[ m  ��\\ �]] & C o n v e r t   :   t o   p . / p p .Y ^_^ k  �G`` aba r  ��cdc m  ��ee �ff  d n     ghg 1  ����
�� 
txdlh 1  ����
�� 
ascrb iji O  �Dklk k  �Cmm non I ��������
�� .miscactvnull��� ��� null��  ��  o pqp r  ��rsr n  ��tut 1  ����
�� 
1717u n  ��vwv 1  ����
�� 
wTxRw o  ������ 0 worddocument wordDocuments o      ���� 0 
findobject 
findObjectq xyx r  ��z{z m  ����
�� boovfals{ n      |}| 1  ����
�� 
1840} o  ������ 0 
findobject 
findObjecty ~��~ X  �C��� l �>���� k  �>�� ��� r  ����� o  ������ 0 citecommand citeCommand� n      ��� 1  ����
�� 
1650� o  ������ 0 
findobject 
findObject� ���� Z  �>������ = ����� c  ����� o  ������ 0 citecommand citeCommand� m  ����
�� 
TEXT� m  ���� ���  [ :  � l ����� k  ��� ��� r  ���� m  ���� ��� 
 [ ,   p .� n      ��� 1  ��
�� 
1650� n  ���� m  ���
�� 
w125� o  ������ 0 
findobject 
findObject� ���� I ����
�� .sWRD1874null���     w124� o  ���� 0 
findobject 
findObject� �����
�� 
5642� m  ��
�� e273� ��  ��  �   normal colon   � ���    n o r m a l   c o l o n� ��� = ��� c  ��� o  ���� 0 citecommand citeCommand� m  ��
�� 
TEXT� m  �� ���  [��  � ���� l !:���� k  !:�� ��� r  !.��� m  !$�� ���  [ ,   p p .� n      ��� 1  )-��
�� 
1650� n  $)��� m  %)��
�� 
w125� o  $%���� 0 
findobject 
findObject� ���� I /:����
�� .sWRD1874null���     w124� o  /0���� 0 
findobject 
findObject� �����
�� 
5642� m  36��
�� e273� ��  ��  �   modifier letter colon   � ��� ,   m o d i f i e r   l e t t e r   c o l o n��  ��  ��  � - ' normal colon and modifier letter colon   � ��� N   n o r m a l   c o l o n   a n d   m o d i f i e r   l e t t e r   c o l o n�� 0 citecommand citeCommand� J  ���� ��� m  ���� ���  [ :  � ���� m  ���� ���  [��  ��  ��  l m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  j ���� L  EG����  ��  _ ��� E  JO��� o  JK���� 0 
menuchoice 
menuChoice� m  KN�� ��� " A d d   L e a d i n g   S p a c e� ��� k  RO�� ��� r  R]��� m  RU�� ���  � n     ��� 1  X\��
�� 
txdl� 1  UX��
�� 
ascr� ��� O  ^L��� k  bK�� ��� I bg������
�� .miscactvnull��� ��� null��  ��  � ��� r  hs��� n  hq��� 1  mq��
�� 
1717� n  hm��� 1  im��
�� 
wTxR� o  hi���� 0 worddocument wordDocument� o      ���� 0 
findobject 
findObject� ��� r  t{��� m  tu��
�� boovfals� n      ��� 1  vz��
�� 
1840� o  uv���� 0 
findobject 
findObject� ���� X  |K����� k  �F�� ��� r  ����� o  ������ 0 citecommand citeCommand� n      ��� 1  ����
�� 
1650� o  ������ 0 
findobject 
findObject� ���� Z  �F ��  = �� c  �� o  ������ 0 citecommand citeCommand m  ����
�� 
TEXT m  �� �  \ c i t e p k  ��		 

 r  �� m  �� �    \ c i t e p n       1  ����
�� 
1650 n  �� m  ����
�� 
w125 o  ������ 0 
findobject 
findObject �� I ����
�� .sWRD1874null���     w124 o  ������ 0 
findobject 
findObject ����
�� 
5642 m  ����
�� e273� ��  ��    = �� c  �� o  ������ 0 citecommand citeCommand m  ����
�� 
TEXT m  �� �  \ c i t e t  !  k  ��"" #$# r  ��%&% m  ��'' �((    \ c i t e t& n      )*) 1  ����
�� 
1650* n  ��+,+ m  ����
�� 
w125, o  ���� 0 
findobject 
findObject$ -�~- I ���}./
�} .sWRD1874null���     w124. o  ���|�| 0 
findobject 
findObject/ �{0�z
�{ 
56420 m  ���y
�y e273� �z  �~  ! 121 = ��343 c  ��565 o  ���x�x 0 citecommand citeCommand6 m  ���w
�w 
TEXT4 m  ��77 �88  \ c i t e a l p2 9:9 k  ;; <=< r  >?> m  @@ �AA    \ c i t e a l p? n      BCB 1  	�v
�v 
1650C n  	DED m  	�u
�u 
w125E o  �t�t 0 
findobject 
findObject= F�sF I �rGH
�r .sWRD1874null���     w124G o  �q�q 0 
findobject 
findObjectH �pI�o
�p 
5642I m  �n
�n e273� �o  �s  : JKJ = &LML c  "NON o  �m�m 0 citecommand citeCommandO m  !�l
�l 
TEXTM m  "%PP �QQ  \ c i t e a l tK R�kR k  )BSS TUT r  )6VWV m  ),XX �YY    \ c i t e a l tW n      Z[Z 1  15�j
�j 
1650[ n  ,1\]\ m  -1�i
�i 
w125] o  ,-�h�h 0 
findobject 
findObjectU ^�g^ I 7B�f_`
�f .sWRD1874null���     w124_ o  78�e�e 0 
findobject 
findObject` �da�c
�d 
5642a m  ;>�b
�b e273� �c  �g  �k  ��  ��  �� 0 citecommand citeCommand� J  �bb cdc m  �ee �ff  \ c i t e pd ghg m  ��ii �jj  \ c i t e th klk m  ��mm �nn  \ c i t e a l pl o�ao m  ��pp �qq  \ c i t e a l t�a  ��  � m  ^_rr�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � s�`s L  MO�_�_  �`  � tut E  RWvwv o  RS�^�^ 0 
menuchoice 
menuChoicew m  SVxx �yy ( R e m o v e   L e a d i n g   S p a c eu z�]z k  ZW{{ |}| r  Ze~~ m  Z]�� ���   n     ��� 1  `d�\
�\ 
txdl� 1  ]`�[
�[ 
ascr} ��� O  fT��� k  jS�� ��� I jo�Z�Y�X
�Z .miscactvnull��� ��� null�Y  �X  � ��� r  p{��� n  py��� 1  uy�W
�W 
1717� n  pu��� 1  qu�V
�V 
wTxR� o  pq�U�U 0 worddocument wordDocument� o      �T�T 0 
findobject 
findObject� ��� r  |���� m  |}�S
�S boovfals� n      ��� 1  ~��R
�R 
1840� o  }~�Q�Q 0 
findobject 
findObject� ��P� X  �S��O�� k  �N�� ��� r  ����� o  ���N�N 0 citecommand citeCommand� n      ��� 1  ���M
�M 
1650� o  ���L�L 0 
findobject 
findObject� ��K� Z  �N����J� = ����� c  ����� o  ���I�I 0 citecommand citeCommand� m  ���H
�H 
TEXT� m  ���� ���    \ c i t e p� k  ���� ��� r  ����� m  ���� ���  \ c i t e p� n      ��� 1  ���G
�G 
1650� n  ����� m  ���F
�F 
w125� o  ���E�E 0 
findobject 
findObject� ��D� I ���C��
�C .sWRD1874null���     w124� o  ���B�B 0 
findobject 
findObject� �A��@
�A 
5642� m  ���?
�? e273� �@  �D  � ��� = ����� c  ����� o  ���>�> 0 citecommand citeCommand� m  ���=
�= 
TEXT� m  ���� ���    \ c i t e t� ��� k  ���� ��� r  ����� m  ���� ���  \ c i t e t� n      ��� 1  ���<
�< 
1650� n  ����� m  ���;
�; 
w125� o  ���:�: 0 
findobject 
findObject� ��9� I ���8��
�8 .sWRD1874null���     w124� o  ���7�7 0 
findobject 
findObject� �6��5
�6 
5642� m  ���4
�4 e273� �5  �9  � ��� = ���� c  ���� o  ���3�3 0 citecommand citeCommand� m  ��2
�2 
TEXT� m  �� ���    \ c i t e a l p� ��� k  	"�� ��� r  	��� m  	�� ���  \ c i t e a l p� n      ��� 1  �1
�1 
1650� n  ��� m  �0
�0 
w125� o  �/�/ 0 
findobject 
findObject� ��.� I "�-��
�- .sWRD1874null���     w124� o  �,�, 0 
findobject 
findObject� �+��*
�+ 
5642� m  �)
�) e273� �*  �.  � ��� = %.��� c  %*��� o  %&�(�( 0 citecommand citeCommand� m  &)�'
�' 
TEXT� m  *-�� ���    \ c i t e a l t� ��&� k  1J�� ��� r  1>��� m  14�� ���  \ c i t e a l t� n      ��� 1  9=�%
�% 
1650� n  49   m  59�$
�$ 
w125 o  45�#�# 0 
findobject 
findObject� �" I ?J�!
�! .sWRD1874null���     w124 o  ?@� �  0 
findobject 
findObject ��
� 
5642 m  CF�
� e273� �  �"  �&  �J  �K  �O 0 citecommand citeCommand� J  ��  m  ��		 �

    \ c i t e p  m  �� �    \ c i t e t  m  �� �    \ c i t e a l p � m  �� �    \ c i t e a l t�  �P  � m  fg�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � � L  UW��  �  �]  �d  �  l \\��   , & restore Word to its original position    � L   r e s t o r e   W o r d   t o   i t s   o r i g i n a l   p o s i t i o n � O  \� Q  `� !"  I ch�#�
� .miscslctnull���    obj # o  cd��  0 selectionrange selectionRange�  ! R      ���
� .ascrerr ****      � ****�  �  " l p�$%&$ I p��'�
� .miscslctnull���    obj ' l p�(��( I p��)*
� .sWRD1430null���     docu) o  pq�� 0 worddocument wordDocument* �+,
� 
5263+ l rw-�
�	- n  rw./. 1  uw�
� 
2905/ 1  ru�
� 
sele�
  �	  , �0�
� 
52640 l x}1��1 n  x}232 1  {}�
� 
29053 1  x{�
� 
sele�  �  �  �  �  �  % > 8 produces error if cursor was at the end of the document   & �44 p   p r o d u c e s   e r r o r   i f   c u r s o r   w a s   a t   t h e   e n d   o f   t h e   d o c u m e n t m  \]55�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  �  � 6� 6 L  ������  �   � 787 l     ��������  ��  ��  8 9:9 l     ��;<��  ;  -   < �==  -: >?> l     ��������  ��  ��  ? @A@ i    BCB I      ��D���� 0 settingsmenu settingsMenuD E��E o      ���� ,0 bibliographysettings bibliographySettings��  ��  C k    ]FF GHG l     ��IJ��  I   global setup   J �KK    g l o b a l   s e t u pH LML O     NON r    
PQP 4   ��R
�� 
docuR m    ���� Q o      ���� "0 bibdeskdocument bibDeskDocumentO m     SS�                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  M TUT l   ��VW��  V "  parse bibliography settings   W �XX 8   p a r s e   b i b l i o g r a p h y   s e t t i n g sU YZY r    [\[ m    ]] �^^  \ n     _`_ 1    ��
�� 
txdl` 1    ��
�� 
ascrZ aba r    ocdc o    ���� ,0 bibliographysettings bibliographySettingsd J      ee fgf o      ���� *0 bibdeskdocumentname bibDeskDocumentNameg hih o      ���� 40 bibliographytemplatepath bibliographyTemplatePathi jkj o      ���� &0 citeptemplatepath citepTemplatePathk lml o      ���� &0 citettemplatepath citetTemplatePathm non o      ���� "0 parenthesisopen parenthesisOpeno pqp o      ���� $0 parenthesisclose parenthesisCloseq rsr o      ���� (0 bibliographysortby bibliographySortBys tut o      ���� &0 bibliographystyle bibliographyStyleu vwv o      ���� 0 italiclatin italicLatinw xyx o      ���� 00 italicpublicationtypes italicPublicationTypesy z{z o      ���� .0 superscriptreferences superscriptReferences{ |}| o      ���� (0 numberedreferences numberedReferences} ~��~ o      ���� "0 numberseperator numberSeperator��  b � l  p p������  � 8 2 parse custom lists from the bibliography settings   � ��� d   p a r s e   c u s t o m   l i s t s   f r o m   t h e   b i b l i o g r a p h y   s e t t i n g s� ��� r   p w��� m   p s�� ���  ,� n     ��� 1   t v��
�� 
txdl� 1   s t��
�� 
ascr� ��� r   x ��� n   x }��� 2  y }��
�� 
citm� o   x y���� 0 italiclatin italicLatin� o      ���� 0 italiclatin italicLatin� ��� r   � ���� n   � ���� 2  � ���
�� 
citm� o   � ����� 00 italicpublicationtypes italicPublicationTypes� o      ���� 00 italicpublicationtypes italicPublicationTypes� ��� r   � ���� m   � ��� ���  � n     ��� 1   � ���
�� 
txdl� 1   � ���
�� 
ascr� ��� l   � �������  �
	-- check to make sure that the template files exist	try		tell application "Finder"			get exists of (POSIX file (bibliographyTemplatePath as string) as alias)			get exists of (POSIX file (citepTemplatePath as string) as alias)			get exists of (POSIX file (citetTemplatePath as string) as alias)		end tell	on error		tell application "Microsoft Word"			display dialog "Template files not found, please delete the bibliography and try again" buttons {"OK"} default button {"OK"}		end tell		return	end try
	   � ���  	 - -   c h e c k   t o   m a k e   s u r e   t h a t   t h e   t e m p l a t e   f i l e s   e x i s t  	 t r y  	 	 t e l l   a p p l i c a t i o n   " F i n d e r "  	 	 	 g e t   e x i s t s   o f   ( P O S I X   f i l e   ( b i b l i o g r a p h y T e m p l a t e P a t h   a s   s t r i n g )   a s   a l i a s )  	 	 	 g e t   e x i s t s   o f   ( P O S I X   f i l e   ( c i t e p T e m p l a t e P a t h   a s   s t r i n g )   a s   a l i a s )  	 	 	 g e t   e x i s t s   o f   ( P O S I X   f i l e   ( c i t e t T e m p l a t e P a t h   a s   s t r i n g )   a s   a l i a s )  	 	 e n d   t e l l  	 o n   e r r o r  	 	 t e l l   a p p l i c a t i o n   " M i c r o s o f t   W o r d "  	 	 	 d i s p l a y   d i a l o g   " T e m p l a t e   f i l e s   n o t   f o u n d ,   p l e a s e   d e l e t e   t h e   b i b l i o g r a p h y   a n d   t r y   a g a i n "   b u t t o n s   { " O K " }   d e f a u l t   b u t t o n   { " O K " }  	 	 e n d   t e l l  	 	 r e t u r n  	 e n d   t r y 
 	� ��� T   �L�� k   �G�� ��� l  � �������  �   parse menu choices   � ��� &   p a r s e   m e n u   c h o i c e s� ��� r   � ���� b   � ���� m   � ��� ��� ( B i b D e s k   D o c u m e n t :   	 	� o   � ����� *0 bibdeskdocumentname bibDeskDocumentName� o      ���� 60 bibdeskdocumentnamechoice bibDeskDocumentNameChoice� ��� r   � ���� m   � ��� ���  /� n     ��� 1   � ���
�� 
txdl� 1   � ���
�� 
ascr� ��� r   � ���� b   � ���� m   � ��� ��� 0 B i b l i o g r a p h y   T e m p l a t e :   	� l  � ������� n   � ���� 4   � ����
�� 
citm� m   � �������� o   � ����� 40 bibliographytemplatepath bibliographyTemplatePath��  ��  � o      ���� @0 bibliographytemplatepathchoice bibliographyTemplatePathChoice� ��� r   � ���� b   � ���� m   � ��� ��� 0 A u t h o r - Y e a r   T e m p l a t e :   	 	� l  � ������� n   � ���� 4   � ����
�� 
citm� m   � �������� o   � ����� &0 citeptemplatepath citepTemplatePath��  ��  � o      ���� 20 citeptemplatepathchoice citepTemplatePathChoice� ��� r   � ���� b   � ���� m   � ��� ��� , Y e a r   O n l y   T e m p l a t e :   	 	� l  � ������� n   � ���� 4   � ����
�� 
citm� m   � �������� o   � ����� &0 citettemplatepath citetTemplatePath��  ��  � o      ���� 20 citettemplatepathchoice citetTemplatePathChoice� ��� r   � ���� m   � ��� ���  � n     ��� 1   � ���
�� 
txdl� 1   � ���
�� 
ascr� ��� r   � ���� b   � ���� m   � ��� ��� ( P a r e n t h e s i s   O p e n :   	 	� o   � ����� "0 parenthesisopen parenthesisOpen� o      ���� .0 parenthesisopenchoice parenthesisOpenChoice� ��� r   � ���� b   � ���� m   � ��� ��� * P a r e n t h e s i s   C l o s e :   	 	� o   � ����� $0 parenthesisclose parenthesisClose� o      ���� 00 parenthesisclosechoice parenthesisCloseChoice� ��� r   � ���� m   � ��� ���  ,� n        1   � ���
�� 
txdl 1   � ���
�� 
ascr�  r   � � b   � � m   � � �		 $ I t a l i c   L a t i n :   	 	 	 	 o   � ����� 0 italiclatin italicLatin o      ���� &0 italiclatinchoice italicLatinChoice 

 r   	 b    m    � 6 I t a l i c   P u b l i c a t i o n   T y p e s :   	 o  ���� 00 italicpublicationtypes italicPublicationTypes o      ���� <0 italicpublicationtypeschoice italicPublicationTypesChoice  r  
 m  
 �   n      1  ��
�� 
txdl 1  ��
�� 
ascr  r   b   m     �!! 0 B i b l i o g r a p h y   S o r t   B y :   	 	 o  ���� (0 bibliographysortby bibliographySortBy o      ���� 40 bibliographysortbychoice bibliographySortByChoice "#" r  %$%$ b  !&'& m  (( �)) , B i b l i o g r a p h y   S t y l e :   	 	' o   ���� &0 bibliographystyle bibliographyStyle% o      ���� 20 bibliographystylechoice bibliographyStyleChoice# *+* r  &/,-, b  &+./. m  &)00 �11 2 S u p e r s c r i p t   r e f e r e n c e s :   	/ o  )*���� .0 superscriptreferences superscriptReferences- o      ���� :0 superscriptreferenceschoice superscriptReferencesChoice+ 232 r  09454 b  05676 m  0388 �99 . N u m b e r e d   r e f e r e n c e s :   	 	7 o  34���� (0 numberedreferences numberedReferences5 o      ���� 40 numberedreferenceschoice numberedReferencesChoice3 :;: r  :C<=< b  :?>?> m  :=@@ �AA ( N u m b e r   S e p e r a t o r :   	 	? o  =>���� "0 numberseperator numberSeperator= o      ���� .0 numberseperatorchoice numberSeperatorChoice; BCB l DD��������  ��  ��  C DED r  DKFGF m  DGHH �II Z B u i l t - I n   Q u i c k   S e t t i n g s :   	 H a r v a r d   R e f e r e n c i n gG o      ���� ,0 builtinpreset1choice builtInPreset1ChoiceE JKJ r  LSLML m  LONN �OO V B u i l t - I n   Q u i c k   S e t t i n g s :   	 N u m b e r e d   E n d n o t e sM o      ���� ,0 builtinpreset2choice builtInPreset2ChoiceK PQP O  T�RSR k  Z�TT UVU I Z���WX
�� .gtqpchltns    @   @ ns  W J  Z�YY Z[Z o  Z[���� 60 bibdeskdocumentnamechoice bibDeskDocumentNameChoice[ \]\ o  [^���� @0 bibliographytemplatepathchoice bibliographyTemplatePathChoice] ^_^ o  ^a���� 20 citeptemplatepathchoice citepTemplatePathChoice_ `a` o  ad���� 20 citettemplatepathchoice citetTemplatePathChoicea bcb o  dg���� .0 parenthesisopenchoice parenthesisOpenChoicec ded o  gj���� 00 parenthesisclosechoice parenthesisCloseChoicee fgf o  jm���� &0 italiclatinchoice italicLatinChoiceg hih o  mp���� <0 italicpublicationtypeschoice italicPublicationTypesChoicei jkj o  ps���� 40 bibliographysortbychoice bibliographySortByChoicek lml o  sv���� 20 bibliographystylechoice bibliographyStyleChoicem non o  vy���� :0 superscriptreferenceschoice superscriptReferencesChoiceo pqp o  y|���� 40 numberedreferenceschoice numberedReferencesChoiceq rsr o  |���� .0 numberseperatorchoice numberSeperatorChoices tut m  �vv �ww   u xyx o  ������ ,0 builtinpreset1choice builtInPreset1Choicey z��z o  ������ ,0 builtinpreset2choice builtInPreset2Choice��  X ��{|
�� 
appr{ m  ��}} �~~ & W o r d S c r i p t   S e t t i n g s| ���
�� 
prmp m  ���� ��� F P l e a s e   c h o o s e   w h i c h   s e t t i n g   t o   e d i t� ����
�� 
cnbt� J  ���� ���� m  ���� ���  E x i t��  � �����
�� 
okbt� J  ���� ���� m  ���� ���  C h o o s e��  ��  V ���� r  ����� 1  ����
�� 
rslt� o      ���� 0 
menuchoice 
menuChoice��  S m  TW���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  Q ��� l ��������  �   save and exit   � ���    s a v e   a n d   e x i t� ��� Z  ��������� = ����� o  ���� 0 
menuchoice 
menuChoice� m  ���~
�~ boovfals�  S  ����  ��  � ��� l ���}���}  �   Document Settings   � ��� $   D o c u m e n t   S e t t i n g s� ��|� Z  �G����{� E  ����� o  ���z�z 0 
menuchoice 
menuChoice� o  ���y�y 60 bibdeskdocumentnamechoice bibDeskDocumentNameChoice� Q  ������ k  ���� ��� O  ����� I ���x��w
�x .sysodlogaskr        TEXT� m  ���� ��� Z O p e n   d e s i r e d   l i b r a r y   i n   B i b D e s k   t h e n   c l i c k   O K�w  � m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��v� O  ����� r  ����� n  ����� 1  ���u
�u 
pnam� o  ���t�t "0 bibdeskdocument bibDeskDocument� o      �s�s *0 bibdeskdocumentname bibDeskDocumentName� m  �����                                                                                  BDSK  alis    &  Macintosh HD                   BD ����BibDesk.app                                                    ����            ����  
 cu             Applications  /:Applications:BibDesk.app/     B i b D e s k . a p p    M a c i n t o s h   H D  Applications/BibDesk.app  / ��  �v  � R      �r�q�p
�r .ascrerr ****      � ****�q  �p  � l ���o���o  �   do nothing   � ���    d o   n o t h i n g� ��� E  ����� o  ���n�n 0 
menuchoice 
menuChoice� o  ���m�m @0 bibliographytemplatepathchoice bibliographyTemplatePathChoice� ��� O   #��� r  "��� n   ��� 1   �l
�l 
psxp� l ��k�j� I �i�h�
�i .sysostdfalis    ��� null�h  � �g��
�g 
prmp� m  
�� ��� F P l e a s e   c h o o s e   b i b l i o g r a p h y   t e m p l a t e� �f��e
�f 
ftyp� J  �� ��� m  �� ���  p u b l i c . r t f� ��d� m  �� ��� , c o m . m i c r o s o f t . w o r d . d o c�d  �e  �k  �j  � o      �c�c 40 bibliographytemplatepath bibliographyTemplatePath� m   ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� E  &-��� o  &)�b�b 0 
menuchoice 
menuChoice� o  ),�a�a 20 citeptemplatepathchoice citepTemplatePathChoice� ��� O  0P��� r  6O��� n  6M��� 1  IM�`
�` 
psxp� l 6I��_�^� I 6I�]�\�
�] .sysostdfalis    ��� null�\  � �[��
�[ 
prmp� m  :=�� ��� T P l e a s e   c h o o s e   a u t h o r - y e a r   t e m p l a t e   ( c i t e p )� �Z��Y
�Z 
ftyp� J  @E�� ��X� m  @C�� ��� " p u b l i c . p l a i n - t e x t�X  �Y  �_  �^  � o      �W�W &0 citeptemplatepath citepTemplatePath� m  03���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� E  SZ��� o  SV�V�V 0 
menuchoice 
menuChoice� o  VY�U�U 20 citettemplatepathchoice citetTemplatePathChoice� ��� O  ]}��� r  c|��� n  cz��� 1  vz�T
�T 
psxp� l cv��S�R� I cv�Q�P 
�Q .sysostdfalis    ��� null�P    �O
�O 
prmp m  gj � P P l e a s e   c h o o s e   y e a r   o n l y   t e m p l a t e   ( c i t e t ) �N�M
�N 
ftyp J  mr �L m  mp �		 " p u b l i c . p l a i n - t e x t�L  �M  �S  �R  � o      �K�K &0 citettemplatepath citetTemplatePath� m  ]`

�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  �  E  �� o  ���J�J 0 
menuchoice 
menuChoice o  ���I�I .0 parenthesisopenchoice parenthesisOpenChoice  O  �� k  ��  I ���H
�H .sysodlogaskr        TEXT m  �� � R S e t   P a r e n t h e s i s   O p e n 
 E x a m p l e :   (   o r   [   o r   < �G
�G 
btns J  �� �F m  �� �  O K�F   �E !
�E 
dflt  m  ��"" �##  O K! �D$�C
�D 
dtxt$ o  ���B�B "0 parenthesisopen parenthesisOpen�C   %�A% r  ��&'& n  ��()( 1  ���@
�@ 
ttxt) 1  ���?
�? 
rslt' o      �>�> "0 parenthesisopen parenthesisOpen�A   m  ��**�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��   +,+ E  ��-.- o  ���=�= 0 
menuchoice 
menuChoice. o  ���<�< 00 parenthesisclosechoice parenthesisCloseChoice, /0/ O  ��121 k  ��33 454 I ���;67
�; .sysodlogaskr        TEXT6 m  ��88 �99 T S e t   P a r e n t h e s i s   C l o s e 
 E x a m p l e :   )   o r   ]   o r   =7 �::;
�: 
btns: J  ��<< =�9= m  ��>> �??  O K�9  ; �8@A
�8 
dflt@ m  ��BB �CC  O KA �7D�6
�7 
dtxtD o  ���5�5 $0 parenthesisclose parenthesisClose�6  5 E�4E r  ��FGF n  ��HIH 1  ���3
�3 
ttxtI 1  ���2
�2 
rsltG o      �1�1 $0 parenthesisclose parenthesisClose�4  2 m  ��JJ�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  0 KLK E  ��MNM o  ���0�0 0 
menuchoice 
menuChoiceN o  ���/�/ &0 italiclatinchoice italicLatinChoiceL OPO k  �BQQ RSR r  ��TUT m  ��VV �WW  ,U n     XYX 1  ���.
�. 
txdlY 1  ���-
�- 
ascrS Z[Z O   :\]\ k  9^^ _`_ I #�,ab
�, .sysodlogaskr        TEXTa m  	cc �dd � S e t   p h r a s e s   t o   b e   i t a l i s i s e d ,   s e p a r a t e d   b y   c o m m a s 
 E x a m p l e :   e t   a l . , i b i d . , c f . , i n t e r   a l i a , c i r c ab �+ef
�+ 
btnse J  gg h�*h m  ii �jj  O K�*  f �)kl
�) 
dfltk m  mm �nn  O Kl �(o�'
�( 
dtxto c  pqp o  �&�& 0 italiclatin italicLatinq m  �%
�% 
TEXT�'  ` rsr r  $-tut n  $+vwv 1  '+�$
�$ 
ttxtw 1  $'�#
�# 
rsltu o      �"�" 0 italiclatin italicLatins x�!x r  .9yzy c  .7{|{ l .3}� �} n  .3~~ 2 /3�
� 
citm o  ./�� 0 italiclatin italicLatin�   �  | m  36�
� 
listz o      �� 0 italiclatin italicLatin�!  ] m   ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  [ ��� r  ;B��� m  ;>�� ���  � n     ��� 1  ?A�
� 
txdl� 1  >?�
� 
ascr�  P ��� E  EL��� o  EH�� 0 
menuchoice 
menuChoice� o  HK�� <0 italicpublicationtypeschoice italicPublicationTypesChoice� ��� k  O��� ��� r  OV��� m  OR�� ���  ,� n     ��� 1  SU�
� 
txdl� 1  RS�
� 
ascr� ��� O  W���� k  ]��� ��� I ]z���
� .sysodlogaskr        TEXT� m  ]`�� ��� � S e t   r e f e r e n c e   t y p e s   t o   b e   i t a l i s i s e d ,   s e p a r a t e d   b y   c o m m a s 
 E x a m p l e :   f i l m , b r o a d c a s t� ���
� 
btns� J  ch�� ��� m  cf�� ���  O K�  � ���
� 
dflt� m  kn�� ���  O K� ���
� 
dtxt� c  qv��� o  qr�� 00 italicpublicationtypes italicPublicationTypes� m  ru�
� 
TEXT�  � ��� r  {���� n  {���� 1  ~��
� 
ttxt� 1  {~�

�
 
rslt� o      �	�	 00 italicpublicationtypes italicPublicationTypes� ��� r  ����� c  ����� l ������ n  ����� 2 ���
� 
citm� o  ���� 00 italicpublicationtypes italicPublicationTypes�  �  � m  ���
� 
list� o      �� 00 italicpublicationtypes italicPublicationTypes�  � m  WZ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� r  ����� m  ���� ���  � n     ��� 1  ��� 
�  
txdl� 1  ����
�� 
ascr�  � ��� E  ����� o  ������ 0 
menuchoice 
menuChoice� o  ������ 40 bibliographysortbychoice bibliographySortByChoice� ��� O  ����� k  ���� ��� I ������
�� .sysodlogaskr        TEXT� m  ���� ��� � S e t   S o r t   B i b l i o g r a p h y   B y   B i b D e s k   F i e l d 
 E x a m p l e :   A u t h o r   o r   C i t e   K e y 
 ( L e a v e   b l a n k   t o   s o r t   i n   o r d e r   o f   a p p e a r a n c e )� ����
�� 
btns� J  ���� ���� m  ���� ���  O K��  � ����
�� 
dflt� m  ���� ���  O K� �����
�� 
dtxt� o  ������ (0 bibliographysortby bibliographySortBy��  � ���� r  ����� n  ����� 1  ����
�� 
ttxt� 1  ����
�� 
rslt� o      ���� (0 bibliographysortby bibliographySortBy��  � m  �����                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  � ��� E  ����� o  ������ 0 
menuchoice 
menuChoice� o  ������ 20 bibliographystylechoice bibliographyStyleChoice� ��� O  ���� k  ��� ��� I ������
�� .sysodlogaskr        TEXT� m  ���� ��� S e t   b i b l i o g r a p h y   s t y l e   i n   M i c r o s o f t   W o r d 
 E x a m p l e :   B i b l i o g r a p h y   o r   L i s t   N u m b e r   o r   N o r m a l 
 ( L e a v e   b l a n k   t o   u s e   f o r m a t t i n g   f r o m   t e m p l a t e   f i l e )� ����
�� 
btns� J  ���� ���� m  ���� ���  O K��  � ����
�� 
dflt� m  ���� ���  O K� �����
�� 
dtxt� o  ������ &0 bibliographystyle bibliographyStyle��  � ���� r  �   n  � 1   ��
�� 
ttxt 1  � ��
�� 
rslt o      ���� &0 bibliographystyle bibliographyStyle��  � m  ���                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  �  E  
 o  
���� 0 
menuchoice 
menuChoice o  ���� :0 superscriptreferenceschoice superscriptReferencesChoice 	
	 O  ? k  >  I 0��
�� .sysodlogaskr        TEXT m   � 6 U s e   s u p e r s c r i p t   r e f e r e n c e s ? ��
�� 
btns J   (  m   # �  N o �� m  #& �  Y e s��   ����
�� 
dflt o  +,���� .0 superscriptreferences superscriptReferences��   �� r  1> !  n  1<"#" 2 8<��
�� 
citm# n  18$%$ 1  48��
�� 
bhit% 1  14��
�� 
rslt! o      ���� .0 superscriptreferences superscriptReferences��   m  &&�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  
 '(' E  BI)*) o  BE���� 0 
menuchoice 
menuChoice* o  EH���� 40 numberedreferenceschoice numberedReferencesChoice( +,+ O  Lw-.- k  Rv// 010 I Rh��23
�� .sysodlogaskr        TEXT2 m  RU44 �55 � U s e   n u m b e r e d   r e f e r e n c e s   i n s t e a d   o f   H a r v a r d   s t y l e ? 
 ( O v e r r i d e s   r e f e r e n c e   t e m p l a t e s   a n d   s o r t i n g )3 ��67
�� 
btns6 J  X`88 9:9 m  X[;; �<<  N o: =��= m  [^>> �??  Y e s��  7 ��@��
�� 
dflt@ o  cd���� (0 numberedreferences numberedReferences��  1 A��A r  ivBCB n  itDED 2 pt��
�� 
citmE n  ipFGF 1  lp��
�� 
bhitG 1  il��
�� 
rsltC o      ���� (0 numberedreferences numberedReferences��  . m  LOHH�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  , IJI E  z�KLK o  z}���� 0 
menuchoice 
menuChoiceL o  }����� .0 numberseperatorchoice numberSeperatorChoiceJ MNM O  ��OPO k  ��QQ RSR I ����TU
�� .sysodlogaskr        TEXTT m  ��VV �WW � S e t   s e p e r a t o r   f o r   n u m b e r e d   r e f e r e n c e s 
 ( O n l y   a p p l i e s   t o   m u l t i p l e   n u m b e r s   i n   a   s i n g l e   f i e l d ) 
 E x a m p l e :   ,   o r   ;   o r   s p a c eU ��XY
�� 
btnsX J  ��ZZ [��[ m  ��\\ �]]  O K��  Y ��^_
�� 
dflt^ m  ��`` �aa  O K_ ��b��
�� 
dtxtb o  ������ "0 numberseperator numberSeperator��  S c��c r  ��ded n  ��fgf 1  ����
�� 
ttxtg 1  ����
�� 
rslte o      ���� "0 numberseperator numberSeperator��  P m  ��hh�                                                                                  MSWD  alis    B  Macintosh HD                   BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    M a c i n t o s h   H D  Applications/Microsoft Word.app   / ��  N iji E  ��klk o  ������ 0 
menuchoice 
menuChoicel o  ������ ,0 builtinpreset1choice builtInPreset1Choicej mnm k  �oo pqp r  ��rsr m  ��tt �uu  (s o      ���� "0 parenthesisopen parenthesisOpenq vwv r  ��xyx m  ��zz �{{  )y o      ���� $0 parenthesisclose parenthesisClosew |}| r  ��~~ m  ���� ���  C i t e   K e y o      ���� (0 bibliographysortby bibliographySortBy} ��� r  ����� m  ���� ���  B i b l i o g r a p h y� o      ���� &0 bibliographystyle bibliographyStyle� ��� r  ����� J  ���� ��� m  ���� ���  e t   a l .� ��� m  ���� ��� 
 i b i d .� ��� m  ���� ���  c f .� ��� m  ���� ���  i n t e r   a l i a� ���� m  ���� ��� 
 c i r c a��  � o      ���� 0 italiclatin italicLatin� ��� r  ����� J  ���� ��� m  ���� ���  f i l m� ���� m  ���� ���  b r o a d c a s t��  � o      ���� 00 italicpublicationtypes italicPublicationTypes� ��� r  ����� m  ���� ���  N o� o      ���� .0 superscriptreferences superscriptReferences� ��� r  ����� m  ���� ���  N o� o      ���� (0 numberedreferences numberedReferences� ���� r  ���� m  ��� ���  � o      ���� "0 numberseperator numberSeperator��  n ��� E  ��� o  	���� 0 
menuchoice 
menuChoice� o  	���� ,0 builtinpreset2choice builtInPreset2Choice� ���� k  C�� ��� r  ��� m  �� ���  � o      ���� "0 parenthesisopen parenthesisOpen� ��� r  ��� m  �� ���  � o      ���� $0 parenthesisclose parenthesisClose� ��� r  !��� m  �� ���  � o      ���� (0 bibliographysortby bibliographySortBy� ��� r  "'��� m  "%�� ���  L i s t   N u m b e r� o      ���� &0 bibliographystyle bibliographyStyle� ��� r  (,��� J  (*����  � o      ���� 0 italiclatin italicLatin� ��� r  -1��� J  -/����  � o      ���� 00 italicpublicationtypes italicPublicationTypes� ��� r  27��� m  25�� ���  Y e s� o      ���� .0 superscriptreferences superscriptReferences� ��� r  8=��� m  8;�� ���  Y e s� o      ���� (0 numberedreferences numberedReferences� ���� r  >C��� m  >A�� ���  ,� o      ���� "0 numberseperator numberSeperator��  ��  �{  �|  � ���� L  M]�� J  M\�� ��� o  MN���� *0 bibdeskdocumentname bibDeskDocumentName� ��� o  NO���� 40 bibliographytemplatepath bibliographyTemplatePath� ��� o  OP���� &0 citeptemplatepath citepTemplatePath� ��� o  PQ���� &0 citettemplatepath citetTemplatePath� � � o  QR���� "0 parenthesisopen parenthesisOpen   o  RS���� $0 parenthesisclose parenthesisClose  o  ST���� (0 bibliographysortby bibliographySortBy  o  TU���� &0 bibliographystyle bibliographyStyle  o  UV���� 0 italiclatin italicLatin 	
	 o  VW���� 00 italicpublicationtypes italicPublicationTypes
  o  WX���� .0 superscriptreferences superscriptReferences  o  XY���� (0 numberedreferences numberedReferences �� o  YZ���� "0 numberseperator numberSeperator��  ��  A �� l     ������  �  �  ��       ���������   ��������~�}�|�{�z�y�x�w�v�u�t� "0 defaultsettings defaultSettings� 40 readbibliographysettings readBibliographySettings� 60 writebibliographysettings writeBibliographySettings� $0 formatreferences formatReferences� (0 unformatreferences unformatReferences� *0 fillemptyreferences fillEmptyReferences� 0 	toolsmenu 	toolsMenu�~ 0 settingsmenu settingsMenu
�} .aevtoappnull  �   � ****�| 0 	countword 	countWord�{ 0 countbibdesk countBibDesk�z  0 selectionstart selectionStart�y 0 selectionend selectionEnd�x 0 
menuchoice 
menuChoice�w ,0 bibliographysettings bibliographySettings�v  �u  �t   �sR�r�q�p�s "0 defaultsettings defaultSettings�r  �q   �o�n�m�l�k�j�i�h�g�f�e�d�c�o *0 bibdeskdocumentname bibDeskDocumentName�n 40 bibliographytemplatepath bibliographyTemplatePath�m &0 citeptemplatepath citepTemplatePath�l &0 citettemplatepath citetTemplatePath�k "0 parenthesisopen parenthesisOpen�j $0 parenthesisclose parenthesisClose�i (0 bibliographysortby bibliographySortBy�h &0 bibliographystyle bibliographyStyle�g 0 italiclatin italicLatin�f 00 italicpublicationtypes italicPublicationTypes�e .0 superscriptreferences superscriptReferences�d (0 numberedreferences numberedReferences�c "0 numberseperator numberSeperator X^djpv�������b������a�b �a �p V�E�O�E�O�E�O�E�O�E�O�E�O�E�O�E�O������vE�O��lvE�Oa E�Oa E�Oa E�O�������������a v �`��_�^ �]�` 40 readbibliographysettings readBibliographySettings�_  �^   �\�[�Z�Y�X�W�V�U�T�S�R�Q�P�O�N�M�L�K�J�I�H�G�\ 0 worddocument wordDocument�[  0 selectionrange selectionRange�Z 0 
fieldcount 
fieldCount�Y 0 currentfield currentField�X 0 	fieldcode 	fieldCode�W ,0 bibliographysettings bibliographySettings�V 0 
findobject 
findObject�U 0 	findcount 	findCount�T *0 bibliographycommand bibliographyCommand�S *0 bibdeskdocumentname bibDeskDocumentName�R 40 bibliographytemplatepath bibliographyTemplatePath�Q &0 citeptemplatepath citepTemplatePath�P &0 citettemplatepath citetTemplatePath�O "0 parenthesisopen parenthesisOpen�N $0 parenthesisclose parenthesisClose�M (0 bibliographysortby bibliographySortBy�L &0 bibliographystyle bibliographyStyle�K 0 italiclatin italicLatin�J 00 italicpublicationtypes italicPublicationTypes�I .0 superscriptreferences superscriptReferences�H (0 numberedreferences numberedReferences�G "0 numberseperator numberSeperator  T�F�E�D�C�B�A�@�?�>�=�<�;�:�9�8�7�6�5Q�4�3�2�1�0�/�.�-�,��+�*�)�(��'��&��%��$�#�"�
"&*.1�!� ���������aglx���������
�F 
1003
�E 
5263
�D 
sele
�C 
2903
�B 
5264
�A 
2905�@ 
�? .sWRD1430null���     docu
�> 
1091
�= 
ascr
�< 
txdl
�; 
w170
�: .corecnte****       ****
�9 
2469
�8 
1650
�7 
wFtP
�6 e183G Q
�5 
cwor
�4 
bool
�3 
cha �2 �1��
�0 
TEXT
�/ 
1717
�. 
1840
�- e265� 
�, 
1862
�+ .miscslctnull���    obj 
�* .sWRD1874null���     w124�)  �(  
�' 
btns
�& 
cbtn
�% 
dflt�$ 
�# .sysodlogaskr        TEXT�" �! 
�  
citm
� 
cobj� � � � 	� 
� � � 
� 
psxf
� 
alis
� .coredoexnull���     ****�]?� !*�,E�O��*�,�,�*�,�,� E�O�*�,FUO���,FO�*��-j E�O _h�j ��/E�O�a ,a ,E�O�a ,a  	 �a l/a a & �[a \[Za \Za 2a &E�OY hO�kE�[OY��O�j  �*�,a ,E�Oe�a ,FOa �a ,FOa �a ,FO��j�j� j  O jE�O h�j !�kE�[OY��W +X " #a $a %a &kva 'a (kva )a *kva + ,OhO�j &*�,a ,a &[a \[Za -\Za 2a &E�Y �j  jvY hY hUO �a .a /a 0a 1a 2a 3a 4a 5a 6a 7a 8a 9a :a ;a <v��,FO�a =-E[a >k/E�Z[a >l/E�Z[a >m/E�Z[a >�/E�Z[a >a ?/E�Z[a >a +/E�Z[a >a @/E�Z[a >a A/E�Z[a >a B/E^ Z[a >a C/E^ Z[a >a D/E^ Z[a >a E/E^ Z[a >a F/E^ Z[a >a </E^ ZW &X " #� a Ga %a Hkva )a Ikv� ,UOhOa J��,FO] a =-E^ O] a =-E^ Oa K��,FO Da L :*a M�a &/a N&j OO*a M�a &/a N&j OO*a M�a &/a N&j OUW &X " #� a Pa %a Qkva )a Rkv� ,UOhO� / 
�j  W X " #��*�,�,�*�,�,� j  Oa S*�,FUO�������] ] ] ] ] ] a Fv �*��!"�� 60 writebibliographysettings writeBibliographySettings� �#� #  �� ,0 bibliographysettings bibliographySettings�  ! ����
�	��������� ������������� ,0 bibliographysettings bibliographySettings� 0 worddocument wordDocument�  0 selectionrange selectionRange�
 "0 bibdeskdocument bibDeskDocument�	 *0 bibdeskdocumentname bibDeskDocumentName� 40 bibliographytemplatepath bibliographyTemplatePath� &0 citeptemplatepath citepTemplatePath� &0 citettemplatepath citetTemplatePath� "0 parenthesisopen parenthesisOpen� $0 parenthesisclose parenthesisClose� (0 bibliographysortby bibliographySortBy� &0 bibliographystyle bibliographyStyle� 0 italiclatin italicLatin�  00 italicpublicationtypes italicPublicationTypes�� .0 superscriptreferences superscriptReferences�� (0 numberedreferences numberedReferences�� "0 numberseperator numberSeperator�� 0 
fieldcount 
fieldCount�� &0 bibliographyfield bibliographyField�� *0 assumedtemplatepath assumedTemplatePath" ^P����������������N��X��b����������������������������������������������������������������2��A��GJ����U����i����������������!#%')+-/13579;=G���
�� 
1003
�� 
5263
�� 
sele
�� 
2903
�� 
5264
�� 
2905�� 
�� .sWRD1430null���     docu
�� 
1091
�� 
docu
�� 
ascr
�� 
txdl
�� 
cobj�� �� �� �� �� 	�� 
�� �� �� 
�� 
citm
�� 
w170
�� .corecnte****       ****
�� 
2469
�� 
1650
�� 
2482
�� .sysodlogaskr        TEXT
�� 
kocl
�� 
insh
�� 
wTxR
�� 
prdt
�� 
wFtP
�� e183G #
�� 
2477
�� .corecrel****      � null��  ��  
�� 
pnam
�� 
prmp
�� 
ftyp
�� .sysostdfalis    ��� null
�� 
psxp����
�� 
TEXT
�� 
psxf
�� 
alis
�� .miscslctnull���    obj ��� !*�,E�O��*�,�,�*�,�,� E�O�*�,FUO� *�k/E�UO���,FO�E[a k/E�Z[a l/E�Z[a m/E�Z[a �/E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E^ ZOa ��,FO�a -E�O�a -E�Oa ��,FO� h�a -j E^ O] j Q Kh] j �a ] /a ,a  ,a ! �a ] /E^ Oe] a ",FOY hO] kE^ [OY��Y hUO] j g C� ;a #j $O*a %a a &�a ',6a (a )a *a "ea +a ,a a  -E^ UW 
X . /jvO�a 0  � 	�a 1,E�UY hO�a 2  &� *a 3a 4a 5a 6a 7lv� 8a 9,E�UY hOa :��,FO�[a \[Zk\Za ;2a <&E^ Oa =��,FOa > K B*a ?] a @%a A%a <&/a B&a 9,E�O*a ?] a C%a D%a <&/a B&a 9,E�W X . /hUO�a E  #� *a 3a Fa 5a Gkv� 8a 9,E�UY hO�a H  #� *a 3a Ia 5a Jkv� 8a 9,E�UY hY hO� �a K��,FOa La M%�%a N%�%a O%�%a P%�%a Q%�%a R%�%a S%�%a T%�%a U%�%a V%�%a W%�%a X%�%a Y%] %a Z%] a ,a  ,FOa [��,FO 
�j \W X . /��*�,�,�*�,�,� j \UO� / 
�j \W X . /��*�,�,�*�,�,� j \Oa ]*�,FUO������������] a v �������$%���� $0 formatreferences formatReferences�� ��&�� &  ���� ,0 bibliographysettings bibliographySettings��  $ 4��������������������������������������������������������� ,0 bibliographysettings bibliographySettings�� 0 worddocument wordDocument��  0 selectionrange selectionRange� "0 bibdeskdocument bibDeskDocument� *0 bibdeskdocumentname bibDeskDocumentName� 40 bibliographytemplatepath bibliographyTemplatePath� &0 citeptemplatepath citepTemplatePath� &0 citettemplatepath citetTemplatePath� "0 parenthesisopen parenthesisOpen� $0 parenthesisclose parenthesisClose� (0 bibliographysortby bibliographySortBy� &0 bibliographystyle bibliographyStyle� 0 italiclatin italicLatin� 00 italicpublicationtypes italicPublicationTypes� .0 superscriptreferences superscriptReferences� (0 numberedreferences numberedReferences� "0 numberseperator numberSeperator� 0 citeptemplate citepTemplate� 0 citettemplate citetTemplate� 0 citekeylist citeKeyList� "0 publicationlist publicationList� "0 ibidcitekeylist ibidCiteKeyList� $0 italicreferences italicReferences� 0 formatrange formatRange� 0 
findobject 
findObject� 0 citecommand citeCommand� 0 currentfield currentField� 0 fieldnumber fieldNumber� 0 	fieldcode 	fieldCode� 0 citekey citeKey� "0 referenceprefix referencePrefix� "0 referencesuffix referenceSuffix� ,0 multiplepublications multiplePublications� $0 multiplecitekeys multipleCiteKeys� 0 eachinteger eachInteger�  0 currentcitekey currentCiteKey� (0 currentpublication currentPublication� 0 ibidreference ibidReference� 0 citekeynumber citeKeyNumber� ,0 italicreferencetitle italicReferenceTitle� &0 citekeynumbertext citeKeyNumberText� 0 templatedtext templatedText� (0 formattedreference formattedReference�  0 italiccombined italicCombined� 0 
eachitalic 
eachItalic� 0 
italicfind 
italicFind� 0 
fieldcount 
fieldCount� &0 bibliographyfield bibliographyField� 0 tempdirectory tempDirectory� 40 bibliographytemplatefile bibliographyTemplateFile� <0 bibliographytemplatefilename bibliographyTemplateFileName� 0 tempfile  % ����������������������~�}�|�{�z�y�x,�w�v�uy�t�s�r�q�p�o�n	7	;	?	C	F�m�l���k��j�i�h�g�f�e�d�c��b	�a�`�_�^	%�]	+�\	0�[�Z	n	r	v	y�Y	�	�	�	�	�	�
�X'�W
Z
`
e�V�U!�T�S�R>Liw���Q��P�O�N�M�L�KJ���J�I�H�G�F�E�D�C��B�A�@�?�>�=��<�;�:�9�8�7�6v
� 
1003
� 
5263
� 
sele
� 
2903
� 
5264
� 
2905� 
� .sWRD1430null���     docu
� 
1091
� 
docu
� 
ascr
� 
txdl
� 
cobj� � � �~ �} 	�| 
�{ �z �y 
�x 
citm
�w .rdwrread****        ****
�v .miscactvnull��� ��� null
�u 
1650
�t 
csen
�s 
cwor
�r .miscslctnull���    obj 
�q 
1717
�p 
1840
�o e265� 
�n 
1862
�m 
kocl
�l .corecnte****       ****
�k .sWRD1874null���     w124
�j 
w170
�i 
insh
�h 
wTxR
�g 
prdt
�f 
wFtP
�e e183G #
�d 
2482
�c 
2477
�b .corecrel****      � null
�a 
cha 
�` 
2469�_  �^  
�] 
btns
�\ 
dflt
�[ .sysodlogaskr        TEXT
�Z e183G Q
�Y 
bool
�X 
bibi'  
�W 
ckey
�V 
type
�U 
titl
�T 
usTx
�S 
for 
�R .BDSKttxtnull���     docu
�Q 
2475
�P 
wFnO
�O 
2115
�N 
w125
�M 
ital
�L 
5642
�K e273� 
�J afdrtemp
�I 
from
�H 
fldu
�G .earsffdralis        afdr
�F 
psxf
�E 
TEXT
�D 
file
�C 
pnam
�B 
by  
�A .BDSKSortnull���     ****
�@ 
rslt
�? 
uset
�> 
to  
�= .BDSKexptnull���     ****
�< 
wSCx
�; 
5015
�: .sWRDwiFlnull��� ��� null
�9 .coredelonull���     ditm
�8 
1695
�7 
w173
�6 
1656��	�� !*�,E�O��*�,�,�*�,�,� E�O�*�,FUO� *�k/E�UO���,FO�E[a k/E�Z[a l/E�Z[a m/E�Z[a �/E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E�Z[a a /E^ ZOa ��,FO�a -E�O�a -E�Oa ��,FO�j E^ O�j E^ OjvE^ OjvE^ OjvE^ OjvE^ O��*j O*�,�,*�,�,  8*�,a ,a    *�,a !-a "-E^ Y �E^ O��j�j� j #Y �E^ O*�,a $,E^ Oe] a %,FOa &] a ',FO �a (a )a *a +a ,a v[a -a l .kh a /] %a 0%] a ,FO] j #O � �h] j 1*�,a ,E^ Oa 2*�,a ,FO*a -a 3a 4*�,a 5,a 6a 7a 8a 9ea :a ;a a  <O��*�,�,k�*�,�,� a 3k/E^ Oa =] [a >\[Zl\Zi2%] a ?,a ,F[OY�rW !X @ Aa Ba Ca Dkva Ea Fkv� GOh[OY�%O=k] a 3-j .kh ] a 3] /E^ O] a ?,a ,E^ O] a 7,a H 	 a Ia Ja Ka L�v] a "l/a M&�] j #Oa Na Olv��,FO] a l/E^ Oa Pa Qlv��,FOa RE^ Oa SE^ O] a -j .a   ] a l/E^ O] a �/E^ Y "] a -j .m  ] a l/E^ Y hO��jvE^  OjvE^ !Oa T��,FO�k] a -j .kh "] a -a ] "/E^ #O�a U-a V[a W,\Z] #81E^ $O] $j .j  X� Pe] a 9,FO] a ?,j #O��*�,�,k�*�,�,k� j #Oa X] #%a Ca Ykva Ea Zkv� GOhUY hO] jv	 ] a i/] # a M& 
eE^ %Y fE^ %O] ] #%E^ O] ] # ] ] #%E^ Y hO 0k] j .kh "] a ] "/] #  ] "E^ &Y h[OY��O]  ] $%E^  O] ] $ ] ] $%E^ Y hO��a U-a V[a W,\Z] 81a [, <�a U-a V[a W,\Z] 81a \,E^ 'O] ] ' ] ] '%E^ Y hY hO] !] &%E^ !O] !E^ ([OY�TUO�a M&f N] a ] H] %e  a ^E^ )Y � *�/a _] a `]  � aE^ )UO�] %] )%] %�%E^ *Y �] a b H] %e  a cE^ )Y � *�/a _] a `]  � aE^ )UO�] %] )%] %�%E^ *Y �] a d D] %e  a eE^ )Y � *�/a _] a `]  � aE^ )UO] ] )%] %E^ *Y O] a f D] %e  a gE^ )Y � *�/a _] a `]  � aE^ )UO] ] )%] %E^ *Y hO] *] a h,a ,FY 3�a M&e  (a iE^ *O] ��,FO�] (%�%] a h,a ,FY hO�f  f] a h,a j,a k,FY �e  e] a h,a j,a k,FY hOf] a 9,FO�] %E^ +O �] +[a -a l .kh ,] *] , wa l��,FO a] +[a -a l .kh -*�,a $,E^ O] -] a ,FO] -] a m,a ,FOe] a m,a j,a n,FO] a oa pl 1[OY��Oa q��,FY h[OY�zY h[OY��O] a 3-j .E^ .O�h] .j ] a 3] ./E^ O] a ?,a ,E^ O] a 7,a H 	 ] a "l/a ra M&e] E^ /O] /j #Oa s Ca ta u*a v,l wE^ 0O*a x�/a y&E^ 1O*a z�/a {,E^ 2O] 0a y&] 2%E^ 3UO� W�a |	 �a M&f a M& ] a }�l ~O_ E^ Y hO�a �*a z] 1a y&/a �*a z] 3/a `] a  �UOf] /a 9,FOa �] /a h,a ,FO*a 4��] /a h,a �,�] /a h,a �,� a �] 3� �Oa s *a z] 3/j �UO�] /a h,a �,FO�a ��/a j,a {,] /a h,a j,a {,FO��] /a h,a �,l�] /a h,a �,� j �OY hO] .kE^ .[OY�IUO� / 
�j #W X @ A��*�,�,�*�,�,� j #Oa �*�,FUO� �5��4�3()�2�5 (0 unformatreferences unformatReferences�4  �3  ( �1�0�/�.�-�,�1 0 worddocument wordDocument�0  0 selectionrange selectionRange�/ 0 formatrange formatRange�. 0 	fieldlist 	fieldList�- 0 currentfield currentField�, 0 	fieldcode 	fieldCode) <��+�*�)�(�'�&�%�$��#��"�!� �������������������"&*.25���AE�����
�	�_�����
�+ 
1003
�* 
5263
�) 
sele
�( 
2903
�' 
5264
�& 
2905�% 
�$ .sWRD1430null���     docu
�# 
1091
�" 
ascr
�! 
txdl
�  .miscactvnull��� ��� null
� 
btns
� 
dflt
� .sysodlogaskr        TEXT�  �  
� 
w170
� .corecnte****       ****
� 
rvse
� 
kocl
� 
cobj
� 
2469
� 
1650
� 
wFtP
� e183G Q� 
� 
cwor
� 
bool
� 
2475
� 
1695
� 
cha � 
�
 
TEXT
�	 .miscslctnull���    obj 
� 
5418
� 
insh
� 
wTxR
� .sWRDwITRnull��� ��� null
� .coredelonull���     ditm�2�� !*�,E�O��*�,�,�*�,�,� E�O�*�,FUO���,FO�$*j O*�,�,*�,�,  ,�E�O �a a a lva a � W 	X  hY �E�O�a -E�O�j j  a a a kva a � OhY hO ��a ,[a a l kh �a  ,a !,E�O�a ",a # 	 #a $a %a &a 'a (a )a *v�a +l/a ,& Y�a - a .�a /,a 0,FY hO�[a 1\[Za 2\Zi2a 3&E�O�j 4O*a 5a 6�%a 7*�,a 8,6� 9O�j :Y h[OY�fUO� / 
�j 4W X  ��*�,�,�*�,�,� j 4Oa ;*�,FU ����*+� � *0 fillemptyreferences fillEmptyReferences�  �  * �������������������������� 0 worddocument wordDocument��  0 selectionrange selectionRange�� 0 
findobject 
findObject�� 0 
thecounter 
theCounter�� (0 emptyreferencetext emptyReferenceText�� .0 emptyreferencecontext emptyReferenceContext�� 0 	citeplist 	citepList�� 0 	citetlist 	citetList�� 0 citealplist citealpList�� 0 citeskip citeSkip�� 0 	theresult 	theResult�� 0 templatedtext templatedText+ L�����8����������������������������������3��������@OSWZ����r��z|������������������������������������D��JM^eh��lo��x���
�� 
ascr
�� 
txdl
�� 
1003
�� 
5263
�� 
sele
�� 
2903
�� 
5264
�� 
2905�� 
�� .sWRD1430null���     docu
�� 
1091
�� .miscslctnull���    obj 
�� 
1717
�� 
1840
�� e265�  
�� 
1862
�� 
1650
�� .sWRD1874null���     w124
�� 
csen
�� 
TEXT
�� 
cha ����
�� 
cobj
�� 
appr
�� 
prmp
�� 
inSL�� 
�� .gtqpchltns    @   @ ns  
�� 
rslt
�� 
docu
�� 
usTx
�� 
for 
�� .BDSKttxtnull���     docu
�� 
dtxt
�� 
dflt
�� .sysodlogaskr        TEXT
�� 
ttxt��  ��  
�� 
btns
�� 
cbtn
�� 
bhit�� $0 menuchoiceformat menuChoiceFormat� ����,FO�;*�,E�O��*�,�,�*�,�,� E�O�*�,FO��j�j� j O*�,�,E�Oe�a ,FOa �a ,FOa �a ,FOjE�O�h�j a �%a %*�,FO��*�,�,k�*�,�,� a ,E�Oa *�,a -a ,a &[a \[Zk\Za 2%a %E�OQa a a  a !�vE[a "k/E�Z[a "l/E�Z[a "m/E�Z[a "�/E�ZO�����va #a $a %a &�%a '%a (�a ) *O_ +E�O�� 骦 )a , *a -k/a .a /a 0*a -k/�,� 1E�UY o�� )a , *a -k/a .a 2a 0*a -k/�,� 1E�UY B�� 1�a 3%E�Oa , *a -k/a .a 4a 0*a -k/�,� 1E�UY �f  Y hOa 5�%a #a 6a 7�a 8a 9a ) :O_ +a ;,E�O�a &��*�,�,k�*�,�,� a ,FO�kE�Y hW X < =iE�OO��*�,�,�*�,�,� j [OY�5O 
�j W X < =hUO�j  � a >a ?a @kva 8a A� :UY U�j N B� :a Ba ?a Ca Dlva Ea Fa 8a Ga ) :O_ +a H,a I 
*j+ JY hUW X < =hY hO� 	a K*�,FUOh �������,-���� 0 	toolsmenu 	toolsMenu��  ��  , ������������ 0 worddocument wordDocument��  0 selectionrange selectionRange�� 0 
menuchoice 
menuChoice�� 0 
findobject 
findObject�� 0 citecommand citeCommand- l�����������������������������������sv�����GN����^f~���������QT'8C\e��������eimp'7@PXx�	�����������
� 
1003
� 
5263
� 
sele
� 
2903
� 
5264
� 
2905� 
� .sWRD1430null���     docu
� .miscactvnull��� ��� null� 
� 
appr
� 
prmp
� 
inSL
� 
cnbt
� 
okbt� 

� .gtqpchltns    @   @ ns  
� 
rslt
� 
ascr
� 
txdl
� 
wTxR
� 
1717
� 
1840
� 
kocl
� 
cobj
� .corecnte****       ****
� 
1650
� 
TEXT
� 
w125
� 
5642
� e273� 
� .sWRD1874null���     w124
� .miscslctnull���    obj �  �  ���� *�,E�O*�,�*�,�,�*�,�,� E�UOihZ� @*j 	O������a va a a a a a a a kva a kva  O_ E�UO�f  Y hO�a  hYۢa  �a  _ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO xa &a 'lv[a (a )l *kh ��a +,FO�a ,&a -  a .�a /,a +,FO�a 0a 1l 2Y +�a ,&a 3  a 4�a /,a +,FO�a 0a 1l 2Y h[OY��UOhY)�a 5 �a 6_ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO xa 7a 8lv[a (a )l *kh ��a +,FO�a ,&a 9  a :�a /,a +,FO�a 0a 1l 2Y +�a ,&a ;  a <�a /,a +,FO�a 0a 1l 2Y h[OY��UOhYw�a = �a >_ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO xa ?a @lv[a (a )l *kh ��a +,FO�a ,&a A  a B�a /,a +,FO�a 0a 1l 2Y +�a ,&a C  a D�a /,a +,FO�a 0a 1l 2Y h[OY��UOhYŢa E �a F_ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO xa Ga Hlv[a (a )l *kh ��a +,FO�a ,&a I  a J�a /,a +,FO�a 0a 1l 2Y +�a ,&a K  a L�a /,a +,FO�a 0a 1l 2Y h[OY��UOhY�a Ma N_ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO �a Oa Pa Qa R�v[a (a )l *kh ��a +,FO�a ,&a S  a T�a /,a +,FO�a 0a 1l 2Y {�a ,&a U  a V�a /,a +,FO�a 0a 1l 2Y S�a ,&a W  a X�a /,a +,FO�a 0a 1l 2Y +�a ,&a Y  a Z�a /,a +,FO�a 0a 1l 2Y h[OY�QUOhY�a [a \_ !a ",FO� �*j 	O�a #,a $,E�Of�a %,FO �a ]a ^a _a `�v[a (a )l *kh ��a +,FO�a ,&a a  a b�a /,a +,FO�a 0a 1l 2Y {�a ,&a c  a d�a /,a +,FO�a 0a 1l 2Y S�a ,&a e  a f�a /,a +,FO�a 0a 1l 2Y +�a ,&a g  a h�a /,a +,FO�a 0a 1l 2Y h[OY�QUOhY hO� ' 
�j iW X j k��*�,�,�*�,�,� j iU[OY��Oh �C��./�� 0 settingsmenu settingsMenu� �0� 0  �� ,0 bibliographysettings bibliographySettings�  . �������������������������~�}�|�{�z�y�x� ,0 bibliographysettings bibliographySettings� "0 bibdeskdocument bibDeskDocument� *0 bibdeskdocumentname bibDeskDocumentName� 40 bibliographytemplatepath bibliographyTemplatePath� &0 citeptemplatepath citepTemplatePath� &0 citettemplatepath citetTemplatePath� "0 parenthesisopen parenthesisOpen� $0 parenthesisclose parenthesisClose� (0 bibliographysortby bibliographySortBy� &0 bibliographystyle bibliographyStyle� 0 italiclatin italicLatin� 00 italicpublicationtypes italicPublicationTypes� .0 superscriptreferences superscriptReferences� (0 numberedreferences numberedReferences� "0 numberseperator numberSeperator� 60 bibdeskdocumentnamechoice bibDeskDocumentNameChoice� @0 bibliographytemplatepathchoice bibliographyTemplatePathChoice� 20 citeptemplatepathchoice citepTemplatePathChoice� 20 citettemplatepathchoice citetTemplatePathChoice� .0 parenthesisopenchoice parenthesisOpenChoice� 00 parenthesisclosechoice parenthesisCloseChoice� &0 italiclatinchoice italicLatinChoice� <0 italicpublicationtypeschoice italicPublicationTypesChoice� 40 bibliographysortbychoice bibliographySortByChoice�~ 20 bibliographystylechoice bibliographyStyleChoice�} :0 superscriptreferenceschoice superscriptReferencesChoice�| 40 numberedreferenceschoice numberedReferencesChoice�{ .0 numberseperatorchoice numberSeperatorChoice�z ,0 builtinpreset1choice builtInPreset1Choice�y ,0 builtinpreset2choice builtInPreset2Choice�x 0 
menuchoice 
menuChoice/ }S�w]�v�u�t�s�r�q�p�o�n�m�l�k�j��i���������� (08@HN�v�h�g}�f��e��d��c�b��a�`�_�^��]���\�[���Z�Y"�X�W8>BVcim�V�U�������������T4;>V\`tz�������������������
�w 
docu
�v 
ascr
�u 
txdl
�t 
cobj�s �r �q �p �o �n 	�m 
�l �k �j 
�i 
citm�h 
�g 
appr
�f 
prmp
�e 
cnbt
�d 
okbt
�c .gtqpchltns    @   @ ns  
�b 
rslt
�a .sysodlogaskr        TEXT
�` 
pnam�_  �^  
�] 
ftyp
�\ .sysostdfalis    ��� null
�[ 
psxp
�Z 
btns
�Y 
dflt
�X 
dtxt
�W 
ttxt
�V 
TEXT
�U 
list
�T 
bhit�^� *�k/E�UO���,FO�E[�k/E�Z[�l/E�Z[�m/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�Z[��/E�ZOa ��,FO�a -E�O�a -E�Oa ��,FO�hZa �%E�Oa ��,FOa �a i/%E^ Oa �a i/%E^ Oa �a i/%E^ Oa ��,FOa �%E^ Oa �%E^ Oa ��,FOa �%E^ Oa �%E^ Oa ��,FOa �%E^ Oa  �%E^ Oa !�%E^ Oa "�%E^ Oa #�%E^ Oa $E^ Oa %E^ Oa & \�] ] ] ] ] ] ] ] ] ] ] ] a '] ] a (va )a *a +a ,a -a .kva /a 0kv� 1O_ 2E^ UO] f  Y hO] � ,  a & 	a 3j 4UO� 	�a 5,E�UW X 6 7hYS] ]  (a & *a +a 8a 9a :a ;lv� <a =,E�UY#] ]  %a & *a +a >a 9a ?kv� <a =,E�UY�] ]  %a & *a +a @a 9a Akv� <a =,E�UY�] ]  /a & %a Ba Ca Dkva Ea Fa G�� 4O_ 2a H,E�UY�] ]  /a & %a Ia Ca Jkva Ea Ka G�� 4O_ 2a H,E�UY[] ]  Oa L��,FOa & 5a Ma Ca Nkva Ea Oa G�a P&� 4O_ 2a H,E�O�a -a Q&E�UOa R��,FY] ]  Oa S��,FOa & 5a Ta Ca Ukva Ea Va G�a P&� 4O_ 2a H,E�O�a -a Q&E�UOa W��,FY�] ]  /a & %a Xa Ca Ykva Ea Za G�� 4O_ 2a H,E�UYv] ]  /a & %a [a Ca \kva Ea ]a G�� 4O_ 2a H,E�UY?] ]  0a & &a ^a Ca _a `lva E�� 4O_ 2a a,a -E�UY] ]  0a & &a ba Ca ca dlva E�� 4O_ 2a a,a -E�UY �] ]  /a & %a ea Ca fkva Ea ga G�� 4O_ 2a H,E�UY �] ]  Ma hE�Oa iE�Oa jE�Oa kE�Oa la ma na oa p�vE�Oa qa rlvE�Oa sE�Oa tE�Oa uE�Y C] ]  8a vE�Oa wE�Oa xE�Oa yE�OjvE�OjvE�Oa zE�Oa {E�Oa |E�Y h[OY�JO��������������v �S1�R�Q23�P
�S .aevtoappnull  �   � ****1 k    [44  p55  �66  �77 88 A99 D:: |;; ��O�O  �R  �Q  2  3 J ��N�M�L�K ��J ��I�H � ��G ��F � ��E ��D�C�B�A ��@�?�>�=�<�;�:�9 ��8�7�6�5�4�3�2�1�0�/QUY]`�.�-d�,h�+l�*r�)w�(�'�&���%��$��#,�";
�N 
cobj�M 0 	countword 	countWord�L 0 countbibdesk countBibDesk
�K 
prcs
�J .coredoexnull���     ****
�I 
docu
�H .corecnte****       ****
�G .miscactvnull��� ��� null
�F 
btns
�E 
dflt�D 
�C .sysodlogaskr        TEXT
�B 
rslt
�A 
bhit
�@ .ascrnoop****      � ****
�? .sysodelanull��� ��� nmbr�>  �=  
�< 
sele
�; 
cpar
�: 
wTxR
�9 
1650
�8 .miscslctnull���    obj 
�7 
2903�6  0 selectionstart selectionStart
�5 
2905�4 0 selectionend selectionEnd�3 40 readbibliographysettings readBibliographySettings�2 ,0 bibliographysettings bibliographySettings�1 "0 defaultsettings defaultSettings�0 60 writebibliographysettings writeBibliographySettings�/ $0 formatreferences formatReferences�. 
�- 
appr
�, 
prmp
�+ 
inSL
�* 
cnbt
�) 
okbt�( 

�' .gtqpchltns    @   @ ns  �& 0 
menuchoice 
menuChoice�% (0 unformatreferences unformatReferences�$ *0 fillemptyreferences fillEmptyReferences�# 0 settingsmenu settingsMenu�" 0 	toolsmenu 	toolsMenu�P\� KjjlvE[�k/E�Z[�l/E�ZO*��/j e  ��-j 	E�Y hO*��/j e  ��-j 	E�Y hUO��lvj O*j O <���a lva a a  O_ a ,a  � *j UOmj Y hW 	X  hY hO� M*a ,a k/a ,a ,a   *a ,a k/a ,j !Y hO*a ,a ",E` #O*a ,a $,E` %UO_ #_ % Z*j+ &O_ E` 'O_ 'jv  3*j+ (O_ E` 'O*_ 'k+ )O_ E` 'O_ 'jv  hY hY hO*_ 'k+ *OhY hO*j Ojna +a ,a -a .a /a 0va 1a 2a 3a 4a 5a 6a 7a 8kva 9a :kva ; <O_ E` =oO_ =f  hY hO_ =a > W*j+ &O_ E` 'O_ 'jv  3*j+ (O_ E` 'O*_ 'k+ )O_ E` 'O_ 'jv  hY hY hO*_ 'k+ *YF_ =a ? 
*j+ @Y4_ =a A 
*j+ BY"_ =a C �*j+ &O_ E` 'O_ 'jv  3*j+ (O_ E` 'O*_ 'k+ )O_ E` 'O_ 'jv  hY hY hO*_ 'k+ DO_ _ ' V_ E` 'O*_ 'k+ )O_ E` 'O_ 'jv  hY hO  � *j Oa Ej UO*_ 'k+ *W 	X  hY hY b_ =a F W a Gj O*j+ @W 	X  hO*j+ HO )*j Oa Ij O*j+ &O_ E` 'O*_ 'k+ *W 	X  hY h� � �  �   �!<�! <  == �>> " F o r m a t   R e f e r e n c e s � ?�  ?  @ABCDEFGHIJKL@ �MM  e L i b r a r y . b i bA �NN � / U s e r s / d a v i d / M E G A / S c r i p t s / W o r d S c r i p t / W o r d S c r i p t   T e m p l a t e s / W o r d S c r i p t   T e m p l a t e   B i b l i o g r a p h y . r t fB �OO � / U s e r s / d a v i d / M E G A / S c r i p t s / W o r d S c r i p t / W o r d S c r i p t   T e m p l a t e s / W o r d S c r i p t   T e m p l a t e   A u t h o r - Y e a r   ( c i t e p ) . t x tC �PP � / U s e r s / d a v i d / M E G A / S c r i p t s / W o r d S c r i p t / W o r d S c r i p t   T e m p l a t e s / W o r d S c r i p t   T e m p l a t e   Y e a r   O n l y   ( c i t e t ) . t x tD �QQ  (E �RR  )F �SS  C i t e   K e yG �TT  B i b l i o g r a p h yH �U� U  VWXYZ�����������V �[[  e t   a l .W �\\ 
 i b i d .X �]]  c f .Y �^^  i n t e r   a l i aZ �__ 
 c i r c a�  �  �  �  �  �  �  �  �  �  �  I �`� `  ab���������
�	����a �cc  f i l mb �dd  b r o a d c a s t�  �  �  �  �  �  �  �  �
  �	  �  �  �  �  J �ee  N oK �ff  N oL �gg  ,�  �  �   ascr  ��ޭ