   M S B A M L   `   `   `  ����@  <AppKD, Version=3.0.0.5, Culture=neutral, PublicKeyToken=nullclr-namespace:AppKDAppKD  5   	       AppKD.Window_PDxF   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   %localclr-namespace:AppKD     5   	     � Title$'  !AppKD | Convertir les fichiers 2D��5   	   $	��260��63   $	��500��6@   $	��260��5	   	   $	��500��6    � 
ResizeMode$ NoResize=�5
   	    � WindowStartupLocation$ CenterScreen=�6   .��5      �  � RowDefinitions 5   
   �� 5      $	��100 �6   �� 5      $��80 �6   �� 5      $��25 �6   .+�5   
   �� $��	10,20,0,0q�6   $��0��6(   .+�5      �� $H��BSélection du dossier où se trouvent les fichiers à convertir : ��6   �� 5       �� Orientation$ 
Horizontal=�6   $��0,10,0,0q�62   .+�5      �� -   5         ��Name$ TextPath��5      $��23��6,   $��12(�68   $��Italic$�6F   $	��400��6Y   $��0,0,0,1q�6e   .��6�   ��5      O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem 5   5   �� 6   6sys&clr-namespace:System;assembly=mscorlib     65    �� 
TargetType" ��6     �� 	Resources 5         CueBannerBrush j� 
AlignmentX	 j� 
AlignmentY
 j� Stretch !� Visual��%�   &	        !� 5   "   $
 Left=�6E   $	 Center=�6W   $

 None=�6k    5   &   �� 5   *   $"��Aucun dossier séléctionné��60   
�������$��12(�6f   $��Italic$�6t   5   #   5       �� Triggers 5      P� 5   "    P� Property � �� �� Empty P� Value# Z 6:   . �5    &   ��  �� Property � �� �� Value# [  6C   5!   #   P� 5"   "    � �� 6:   Z� . �5#   &   ��  � ��# [  6C   5$   #   P� 5%   "    �@J�IsKeyboardFocused . . �5&   &   ��  � ��
 �����6C   5'   #   .�5)      ��  � �� 5*   "   �� 5+   &    �� 
TargetType" ��66   .R�5,   *   ��  ��Name$ border��61   #��z��6A   #��z��6m   #��z��6�    J� SnapsToDevicePixels . ._�5-   .   ��  ��Name$ PART_ContentHost��6;   |�.  $��Hidden=�6g   $��Hidden=�6�   5.   +    �� Triggers 5/   *   P� 50   .    � {� .  . �51   2   ��  �� 
TargetName$ border��6L    �@J�Opacity$
 0.56Q�6`   52   /   P� 53   .    �@J�	IsFocused . . �54   2   �� $ border��6P    � �� �#31465a6d   55   /   57   '   59      5:      .��5;   %   5<      �� 5=      -   5        $��10,0,0,0q�5=      
�������$��0q�6>   $��25��6R   $��25��6^   .��5>      �� $��/icons/Open.png��6   5?      5@      5A      �� 5C   
   $��1��6   $��10q�6"   .+�5D      �� -   5        $
��Cpdf��5D      $	��PDF��6#    d� 	IsChecked$
 TrueY�61   �� 5E      -   5        $
��Cdxf��5E      $	��DXF��6#   �� 5F      -   5        $
��Cdwg��5F      $	��DWG��6#   5G      �� 5I   
   $��2��6   $��	260,0,0,0q�6"   $ 
Horizontal=�65   .+�5J      �� -   5        $��OKButton��5J      $��OK��6<   $	��100��6I   �� 5K      -   5        $��CancelButton��5K      $��Annuler��6D   $��10,0,0,0q�6V   $	��100��6h   5L      5M      5O      