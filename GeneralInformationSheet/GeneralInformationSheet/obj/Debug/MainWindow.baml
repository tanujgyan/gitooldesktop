   M S B A M L   `   `   `  ����R  NGeneralInformationSheet, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:GeneralInformationSheetGeneralInformationSheet  5   	   (    "GeneralInformationSheet.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   7local%clr-namespace:GeneralInformationSheet     5   	     � Title$  
MainWindow��5   	   $
��1000��6   $
��1000��6*   .��5	      �  � ColumnDefinitions 5
   
   �� 5      $
��353* �6   �� 5      $	��19* �6   .+�5   
   �� -   5         ��Name$ textBox��5      $
��Left=�6#   $��23��6>   $��	93,32,0,0q�6J    �� TextWrapping$
 Wrap=�6]   $��TextBox��6q   $	��Top=�6�   $	��656��6�   �� 5   
   -   5         ��Name$ 	Data_File��5      $��	Data File��6#   $
��Left=�67   $��	32,28,0,0q�6R   $	��Top=�6e   �� 5   
   -   5         ��Name$ 	btnBrowse��5      $��Browse��6$   $
��Left=�65   $��
768,32,0,0q�6P   $	��Top=�6d   $��75��6|    J� RenderTransformOrigin$ -0.26,0.649&�6�   6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ dvPreviewData��5      $
��Left=�6*   $	��332��6E   $��	32,80,0,0q�6R   $	��Top=�6e   $	��900��6}   $��2��6�   ��6�   �� *Data+6�   �� 5   
   -   5        $ label��5      $��Output File Location��6   $
��Left=�6>   $��
32,459,0,0q�6Y   $	��Top=�6m   �� 5   
   -   5        $ textBox1��5      $
��Left=�6$   $��23��6?   $��167,463,0,0q�6K   $
 Wrap=�6`   $��TextBox��6t   $	��Top=�6�   $	��720��6�   �� 5   
   -   5        $ button��5      $��Generate GI Sheets��6!   $
��Left=�6>   $��167,535,0,0q�6Y   $	��Top=�6n   $	��133��6�   $��39��6�   6�   � 5   
   -   5         �Name$ 
pbDataLoad��5      $
��Left=�6*   $��10��6E   $��
350,60,0,0q�6Q   $	��Top=�6e   $	��100��6}   	 � Minimum$	 0Q�6�   
 � Maximum$	
 100Q�6�   $y�Visible=�6�   5      5      