   M S B A M L   `   `   `  ����@  <AppKD, Version=3.0.0.5, Culture=neutral, PublicKeyToken=nullclr-namespace:AppKDAppKD  5   	       AppKD.Window_UnitsDraw   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   %localclr-namespace:AppKD     5   	     � Title$3  -AppKD | Convertir les unités d'un CATDrawing��5   	   $	��200��6>   $	��500��6K   $	��200��5	   	   $	��500��6    � 
ResizeMode$ NoResize=�5
   	   .��5      �  � RowDefinitions 5   
   �� 5      $��50 �6   �� 5      $��80 �6   �� 5      $��25 �6   .+�5   
   �� $��	10,20,0,0q�6   $��0��6(   .+�5      �� $��Choix de la conversion : ��6   5      �� 5   
   $��1��6   $��10q�6"   .+�5      	� -   5        $	��Cmm��5      $��
inch => mm��6%    d� 	IsChecked$
 TrueY�6:   	� 5      -   5        $��Cinch��5      $��
mm => inch��6'   	� 5      -   5        $��C_all��5      $��Afficher les 2��6'   5      �� 5   
   $��2��6   $��	260,0,0,0q�6"    �� Orientation$ 
Horizontal=�65   .+�5      �� -   5        $��OKButton��5      $��OK��6<   $	��100��6I   �� 5      -   5        $��CancelButton��5      $��Annuler��6D   $��10,0,0,0q�6V   $	��100��6h   5      5      5!      