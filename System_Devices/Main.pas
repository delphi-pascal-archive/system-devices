//=================================================//
//             Nicolas Paglieri (ni69)             //
//                  www.ni69.info                  //
//                & www.delphifr.com               //
//=================================================//
// Merci � DelphiProg pour son aide pr�cieuse ! ;) //
//=================================================//

unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Registry, Grids, ImgList, ShellAPI;

type
  TForm1 = class(TForm)
    EnumerateDevicesBtn: TButton;
    GroupBox1: TGroupBox;
    PropertiesStringGrid: TStringGrid;
    GroupBox2: TGroupBox;
    DevicesTreeView: TTreeView;
    DescTreeView: TTreeView;
    ImageList1: TImageList;
    function Translate(English : string): string;
    procedure EnumerateDevicesBtnClick(Sender: TObject);
    procedure DevicesTreeViewAddition(Sender: TObject; Node: TTreeNode);
    procedure DevicesTreeViewChange(Sender: TObject; Node: TTreeNode);
    procedure FormCreate(Sender: TObject);
  private
    { D�clarations priv�es }
  public
    { D�clarations publiques }
  end;

var
  Form1: TForm1;
  DescList : array of array of String;

implementation

{$R *.dfm}

{$R WindowsXP.res} // Impl�mentation du Style XP dans l'application

//============================================================================//
// Fonction de traduction en fran�ais des noms anglais des cat�gories de p�riph�riques
// On ajoute ici l'index de l'ic�ne de cat�gorie apr�s un # pour g�rer l'affichage
//============================================================================//
function TForm1.Translate(English: string): string;
begin
 if English = 'CDROM' then result := 'Lecteurs de CD-ROM/DVD-ROM#09'
 else if English = 'Computer' then result := 'Ordinateur#14'
 else if English = 'DiskDrive' then result := 'Lecteurs de disque#10'
 else if English = 'Display' then result := 'Cartes Graphiques#01'
 else if English = 'fdc' then result := 'Contr�leur de lecteur de disquettes#04'
 else if English = 'FloppyDisk' then result := 'Lecteurs de disquettes#11'
 else if English = 'hdc' then result := 'Contr�leurs ATA/ATAPI IDE#05'
 else if English = 'Image' then result := 'P�riph�riques d''image#15'
 else if English = 'Keyboard' then result := 'Claviers#03'
 else if English = 'LegacyDriver' then result := 'Pilotes non Plug-and-Play#17'
 else if English = 'MEDIA' then result := 'Contr�leurs audio, vid�o et jeu#06'
 else if English = 'Modem' then result := 'Modems#12'
 else if English = 'Monitor' then result := 'Moniteurs#13'
 else if English = 'Mouse' then result := 'Souris et autres p�riph�riques de pointage#20'
 else if English = 'Net' then result := 'Cartes r�seau#02'
 else if English = 'NtApm' then result := 'Prise en charge NT APM/h�rit�#19'
 else if English = 'Ports' then result := 'Ports (COM et LPT)#18'
 else if English = 'Printer' then result := 'Imprimantes#08'
 else if English = 'System' then result := 'P�riph�riques Syst�me#14'
 else if English = 'USB' then result := 'Contr�leurs de bus USB#07'
 else if English = 'Volume' then result := 'Volumes de stockage#21'
 else result := English+'#22'; // P�riph�rique inconnu
end;
//============================================================================//




//============================================================================//
//          PROCEDURE D'ENUMERATION DES PERIPHERIQUES SUR WINDOWS XP          //
//============================================================================//
procedure TForm1.EnumerateDevicesBtnClick(Sender: TObject);
var
 CategoriesList, SubCatList, SubSubCatList, FinalList : TStringList;
 nb, nb2, nb3: integer;
 Reg1, Reg2, Reg3 : TRegistry;
 HasBeenFound : boolean;
 line : string;
begin

 // On �vite deux �num�rations simultan�es qui entraineraient des probl�mes d'affichage...
 EnumerateDevicesBtn.Enabled := false;

 CategoriesList := TStringList.Create; // Liste des cat�gories principales du registre
 SubCatList := TStringList.Create; // Liste interm�diaire
 SubSubCatList := TStringList.Create; // Liste interm�diaire
 FinalList := TStringList.Create; // Liste finale comprenant les p�riph�riques

 // On cr�e les objets TRegistry qui serviront � parcourir l'arborescence
 Reg1 := TRegistry.Create;
 Reg2 := TRegistry.Create;
 Reg3 := TRegistry.Create;

 try
  // D�finition de la cl� racine
  Reg1.RootKey := HKEY_LOCAL_MACHINE;
  Reg2.RootKey := HKEY_LOCAL_MACHINE;
  Reg3.RootKey := HKEY_LOCAL_MACHINE;

  //----------------------------------------------------------------------------------------
  // 1�re ETAPE : ENUMARTION DES CATEGORIES DU REGISTRE (diff�rentes des cat�gories finales)
  with TRegistry.Create do try
   RootKey := HKEY_LOCAL_MACHINE;

   //! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !
   // IMPORTANT : DROITS D'ACCES
   // On ouvre les cl�s en lecture seule avec OpenKeyReadOnly
   // car on dispose de la valeur de s�curit� d'acc�s KEY_READ.
   // En effet, seul SYSTEM a le droit d'ouvrir cette cl� en �criture en temps normal.
   //! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !

   OpenKeyReadOnly('SYSTEM\CurrentControlSet\Enum');
   GetKeyNames(CategoriesList); // R�cup�ration des cat�gories
   CloseKey;
  finally
   free;
  end;

   //-----------------------------------------------------------------------------------------------------------------------------------
   // 2eme ETAPE : PARCOURS DE L'ARBORESCENCE DU REGISTRE (les cl�s contenant les informations utiles sont contenues dans d'autres cl�s)
   for nb := 0 to CategoriesList.Count-1 do begin
    Reg1.OpenKeyReadOnly('SYSTEM\CurrentControlSet\Enum\'+CategoriesList[nb]);
     Reg1.GetKeyNames(SubCatList);
     Reg1.CloseKey;
     for nb2 := 0 to SubCatList.Count-1 do begin
       Reg2.OpenKeyReadOnly('SYSTEM\CurrentControlSet\Enum\'+CategoriesList[nb]+'\'+SubCatList[nb2]);
       Reg2.GetKeyNames(SubSubCatList);
       Reg2.CloseKey;
       for nb3 := 0 to SubSubCatList.Count-1 do begin
        Reg3.OpenKeyReadOnly('SYSTEM\CurrentControlSet\Enum\'+CategoriesList[nb]+'\'+SubCatList[nb2]+'\'+SubSubCatList[nb3]);

        // Si on ne dispose ni du type de p�riph�rique, ni de son nom,
        // Ou alors si le p�riph�rique n'est plus disponible (si la cl� "Control" n'est pas pr�sente), on ne l'ajoute pas
        if ((Reg3.ReadString('Class')='') and (Reg3.ReadString('DeviceDesc')='')) or (not Reg3.KeyExists('Control'))  then begin
         Reg3.CloseKey;
         continue;
        // Si il s'agit d'un lecteur CD, d'un disque dur ou d'un port (COM ou LPT), on remplace la description du p�riph�rique par un nom plus parlant
        end else if (Reg3.ReadString('Class')='CDROM') or (Reg3.ReadString('Class')='DiskDrive') or (Reg3.ReadString('Class')='Ports') then
         line := Translate(Reg3.ReadString('Class'))+'|'+Reg3.ReadString('FriendlyName')
        else line := Translate(Reg3.ReadString('Class'))+'|'+Reg3.ReadString('DeviceDesc');
        // Ajout des informations si elles sont pr�sentes dans le registre
         if Reg3.ValueExists('DeviceDesc') then Line := Line + '�Description@'+Reg3.ReadString('DeviceDesc');
         if Reg3.ValueExists('FriendlyName') then Line := Line + '�FriendlyName@'+Reg3.ReadString('FriendlyName');
         if Reg3.ValueExists('Mfg') then Line := Line + '�Fabriquant@'+Reg3.ReadString('Mfg');
         if Reg3.ValueExists('Service') then Line := Line + '�Service@'+Reg3.ReadString('Service');
         if Reg3.ValueExists('LocationInformation') then Line := Line + '�Emplacement@'+Reg3.ReadString('LocationInformation');
         if Reg3.ValueExists('Class') then Line := Line + '�Enum�rateur@'+Reg3.ReadString('Class');
        FinalList.Add(line);
        Reg3.CloseKey;
       end;
     end;
   end;

  // On trie la liste des p�riph�riques par ordre alphab�tique
  FinalList.Sort;
  line := '';

  // On vide les TreeViews
  DevicesTreeView.Items.Clear;
  DescTreeView.Items.Clear;

  // 3�me ETAPE : ON REMPLIT LES TREEVIEWS AVEC LA LISTE DES PERIPHERIQUES ET LES INFORMATIONS QUE L'ON CLASSE AU PASSAGE DANS DIFFERENTES CATEGORIES...
  for nb := 0 to FinalList.Count-1 do begin // On parcours tous les p�riph�riques
   HasBeenFound := false; // Variable qui permet de savoir si la cat�gorie existe d�j� dans le TreeView ou si il faut la cr�er
   for nb2 := 0 to DevicesTreeView.Items.Count-1 do begin // On parcours tous les noeuds
    HasBeenFound := ((DevicesTreeView.Items[nb2].Text = Copy(FinalList[nb],1,Pos('#',FinalList[nb])-1)) and (DevicesTreeView.Items[nb2].Level = 0));
    if HasBeenFound then break;
   end;
   if HasBeenFound then begin
     // Si le noeud parent de cat�gorie existe d�j�, on ne fait qu'inclure le p�riph�rique dans cette branche :
     DevicesTreeView.Items.AddChild(DevicesTreeView.Items[nb2],Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Pos('�',FinalList[nb])-Pos('|',FinalList[nb])-1));
     // On cr�e le noeud �quivalent dans le treeview d'infos
     DescTreeView.Items.AddChild(DescTreeView.Items[nb2],Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Length(FinalList[nb])));
   end else if not HasBeenFound then begin
     // Sinon, on cr�e la branche et on y place le p�riph�rique : [ici, le noeud parent est cr�� dans la fonction d'ajout de l'enfant]
     DevicesTreeView.Items.AddChild(  DevicesTreeView.Items.Add(nil, Copy(FinalList[nb],1,Pos('|',FinalList[nb])-1)) , Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Pos('�',FinalList[nb])-Pos('|',FinalList[nb])-1));
//                                   \_________________________cr�ation du noeud parent____________________________/
     // On cr�e le noeud �quivalent dans le treeview d'infos
     DescTreeView.Items.AddChild(     DescTreeView.Items.Add(nil, Copy(FinalList[nb],1,Pos('|',FinalList[nb])-1))   ,Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Length(FinalList[nb])));
//                                   \________________________cr�ation du noeud parent___________________________/
   end;

   // On �vite le blocage de l'application
   Application.ProcessMessages;
  end;

 finally
  // On lib�re les �l�ments cr��s au d�part
  Reg3.Free;
  Reg2.Free;
  Reg1.Free;
  FinalList.Free;
  SubSubCatList.Free;
  SubCatList.Free;
  CategoriesList.Free;
  // On r�autorise l'�num�ration (voir ligne 94)
  EnumerateDevicesBtn.Enabled := true;
 end;
end;
//============================================================================//




//============================================================================//
// Attribution des images de cat�gories pour les noeuds racines et les enfants
//============================================================================//
procedure TForm1.DevicesTreeViewAddition(Sender: TObject; Node: TTreeNode);
begin
 if Node.Level = 0 then begin
  Node.ImageIndex := StrToInt(Copy(Node.Text,Pos('#',Node.Text)+1,2));
  Node.Text := Copy(Node.Text,1,Pos('#',Node.Text)-1);
 end else Node.ImageIndex := Node.Parent.ImageIndex;
end;
//============================================================================//




//============================================================================//
//                  PROCEDURE D'AFFICHAGE DES INFORMATIONS
//============================================================================//
procedure TForm1.DevicesTreeViewChange(Sender: TObject; Node: TTreeNode);
var
 i : integer;
 line : string;
begin

 // On affiche le nom de la cat�gorie dans la partie informations si aucun p�riph�rique particulier n'est s�lectionn�
 if DevicesTreeView.Selected.Level = 0 then begin
   PropertiesStringGrid.RowCount := 1;
   PropertiesStringGrid.Cells[0,0] := 'Category';
   PropertiesStringGrid.Cells[1,0] := DevicesTreeView.Selected.Text;
   exit;
 end;

 // On recherche le noeud correspondant dans le treeview d'informations
 Screen.Cursor := crHourGlass;
 Application.ProcessMessages;
 DescTreeView.Selected := DescTreeView.Items.GetFirstNode;
 while DescTreeView.Selected.AbsoluteIndex <> DevicesTreeView.Selected.AbsoluteIndex do begin
  DescTreeView.Selected := DescTreeView.Selected.GetNext;
  Application.ProcessMessages
 end;
 Screen.Cursor := crDefault;

 // Formatage pr�alable de la ligne (on d�coupe le d�but inutile)
 line := Copy(DescTreeView.Selected.Text,Pos('�', DescTreeView.Selected.Text),Length(DescTreeView.Selected.Text));

 // Et on remplit la StringGrid avec les informations
 i := 0;
 while Pos('�',line)<>0 do begin
  inc(i);
   PropertiesStringGrid.RowCount := i;
   PropertiesStringGrid.Cells[0,i-1] := Copy(line, 2, Pos('@', line)-2);
  if Pos('�',Copy(line,2,Length(line))) = 0 then begin
   PropertiesStringGrid.Cells[1,i-1] := Copy(line,Pos('@', line)+1,Length(line));
   line := '';
  end else begin
   PropertiesStringGrid.Cells[1,i-1] := Copy(line,Pos('@', line)+1,Pos('�',Copy(line,2,Length(line)))-Pos('@', line));
  end;
  // On supprime de la variable line l'information ajout�e
  line := Copy(line,Pos('�',Copy(line,2,Length(line)))+1,Length(line));
 end;
end;       

procedure TForm1.FormCreate(Sender: TObject);
begin
 PropertiesStringGrid.Cells[0,0] := 'device';
 PropertiesStringGrid.Cells[1,0] := 'properties';
end;

end.
