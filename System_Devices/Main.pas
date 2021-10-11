//=================================================//
//             Nicolas Paglieri (ni69)             //
//                  www.ni69.info                  //
//                & www.delphifr.com               //
//=================================================//
// Merci à DelphiProg pour son aide précieuse ! ;) //
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
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Form1: TForm1;
  DescList : array of array of String;

implementation

{$R *.dfm}

{$R WindowsXP.res} // Implémentation du Style XP dans l'application

//============================================================================//
// Fonction de traduction en français des noms anglais des catégories de périphériques
// On ajoute ici l'index de l'icône de catégorie après un # pour gérer l'affichage
//============================================================================//
function TForm1.Translate(English: string): string;
begin
 if English = 'CDROM' then result := 'Lecteurs de CD-ROM/DVD-ROM#09'
 else if English = 'Computer' then result := 'Ordinateur#14'
 else if English = 'DiskDrive' then result := 'Lecteurs de disque#10'
 else if English = 'Display' then result := 'Cartes Graphiques#01'
 else if English = 'fdc' then result := 'Contrôleur de lecteur de disquettes#04'
 else if English = 'FloppyDisk' then result := 'Lecteurs de disquettes#11'
 else if English = 'hdc' then result := 'Contrôleurs ATA/ATAPI IDE#05'
 else if English = 'Image' then result := 'Périphériques d''image#15'
 else if English = 'Keyboard' then result := 'Claviers#03'
 else if English = 'LegacyDriver' then result := 'Pilotes non Plug-and-Play#17'
 else if English = 'MEDIA' then result := 'Contrôleurs audio, vidéo et jeu#06'
 else if English = 'Modem' then result := 'Modems#12'
 else if English = 'Monitor' then result := 'Moniteurs#13'
 else if English = 'Mouse' then result := 'Souris et autres périphériques de pointage#20'
 else if English = 'Net' then result := 'Cartes réseau#02'
 else if English = 'NtApm' then result := 'Prise en charge NT APM/hérité#19'
 else if English = 'Ports' then result := 'Ports (COM et LPT)#18'
 else if English = 'Printer' then result := 'Imprimantes#08'
 else if English = 'System' then result := 'Périphériques Système#14'
 else if English = 'USB' then result := 'Contrôleurs de bus USB#07'
 else if English = 'Volume' then result := 'Volumes de stockage#21'
 else result := English+'#22'; // Périphérique inconnu
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

 // On évite deux énumérations simultanées qui entraineraient des problèmes d'affichage...
 EnumerateDevicesBtn.Enabled := false;

 CategoriesList := TStringList.Create; // Liste des catégories principales du registre
 SubCatList := TStringList.Create; // Liste intermédiaire
 SubSubCatList := TStringList.Create; // Liste intermédiaire
 FinalList := TStringList.Create; // Liste finale comprenant les périphériques

 // On crée les objets TRegistry qui serviront à parcourir l'arborescence
 Reg1 := TRegistry.Create;
 Reg2 := TRegistry.Create;
 Reg3 := TRegistry.Create;

 try
  // Définition de la clé racine
  Reg1.RootKey := HKEY_LOCAL_MACHINE;
  Reg2.RootKey := HKEY_LOCAL_MACHINE;
  Reg3.RootKey := HKEY_LOCAL_MACHINE;

  //----------------------------------------------------------------------------------------
  // 1ère ETAPE : ENUMARTION DES CATEGORIES DU REGISTRE (différentes des catégories finales)
  with TRegistry.Create do try
   RootKey := HKEY_LOCAL_MACHINE;

   //! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !
   // IMPORTANT : DROITS D'ACCES
   // On ouvre les clés en lecture seule avec OpenKeyReadOnly
   // car on dispose de la valeur de sécurité d'accès KEY_READ.
   // En effet, seul SYSTEM a le droit d'ouvrir cette clé en écriture en temps normal.
   //! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !

   OpenKeyReadOnly('SYSTEM\CurrentControlSet\Enum');
   GetKeyNames(CategoriesList); // Récupération des catégories
   CloseKey;
  finally
   free;
  end;

   //-----------------------------------------------------------------------------------------------------------------------------------
   // 2eme ETAPE : PARCOURS DE L'ARBORESCENCE DU REGISTRE (les clés contenant les informations utiles sont contenues dans d'autres clés)
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

        // Si on ne dispose ni du type de périphérique, ni de son nom,
        // Ou alors si le périphérique n'est plus disponible (si la clé "Control" n'est pas présente), on ne l'ajoute pas
        if ((Reg3.ReadString('Class')='') and (Reg3.ReadString('DeviceDesc')='')) or (not Reg3.KeyExists('Control'))  then begin
         Reg3.CloseKey;
         continue;
        // Si il s'agit d'un lecteur CD, d'un disque dur ou d'un port (COM ou LPT), on remplace la description du périphérique par un nom plus parlant
        end else if (Reg3.ReadString('Class')='CDROM') or (Reg3.ReadString('Class')='DiskDrive') or (Reg3.ReadString('Class')='Ports') then
         line := Translate(Reg3.ReadString('Class'))+'|'+Reg3.ReadString('FriendlyName')
        else line := Translate(Reg3.ReadString('Class'))+'|'+Reg3.ReadString('DeviceDesc');
        // Ajout des informations si elles sont présentes dans le registre
         if Reg3.ValueExists('DeviceDesc') then Line := Line + '§Description@'+Reg3.ReadString('DeviceDesc');
         if Reg3.ValueExists('FriendlyName') then Line := Line + '§FriendlyName@'+Reg3.ReadString('FriendlyName');
         if Reg3.ValueExists('Mfg') then Line := Line + '§Fabriquant@'+Reg3.ReadString('Mfg');
         if Reg3.ValueExists('Service') then Line := Line + '§Service@'+Reg3.ReadString('Service');
         if Reg3.ValueExists('LocationInformation') then Line := Line + '§Emplacement@'+Reg3.ReadString('LocationInformation');
         if Reg3.ValueExists('Class') then Line := Line + '§Enumérateur@'+Reg3.ReadString('Class');
        FinalList.Add(line);
        Reg3.CloseKey;
       end;
     end;
   end;

  // On trie la liste des périphériques par ordre alphabétique
  FinalList.Sort;
  line := '';

  // On vide les TreeViews
  DevicesTreeView.Items.Clear;
  DescTreeView.Items.Clear;

  // 3ème ETAPE : ON REMPLIT LES TREEVIEWS AVEC LA LISTE DES PERIPHERIQUES ET LES INFORMATIONS QUE L'ON CLASSE AU PASSAGE DANS DIFFERENTES CATEGORIES...
  for nb := 0 to FinalList.Count-1 do begin // On parcours tous les périphériques
   HasBeenFound := false; // Variable qui permet de savoir si la catégorie existe déjà dans le TreeView ou si il faut la créer
   for nb2 := 0 to DevicesTreeView.Items.Count-1 do begin // On parcours tous les noeuds
    HasBeenFound := ((DevicesTreeView.Items[nb2].Text = Copy(FinalList[nb],1,Pos('#',FinalList[nb])-1)) and (DevicesTreeView.Items[nb2].Level = 0));
    if HasBeenFound then break;
   end;
   if HasBeenFound then begin
     // Si le noeud parent de catégorie existe déjà, on ne fait qu'inclure le périphérique dans cette branche :
     DevicesTreeView.Items.AddChild(DevicesTreeView.Items[nb2],Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Pos('§',FinalList[nb])-Pos('|',FinalList[nb])-1));
     // On crée le noeud équivalent dans le treeview d'infos
     DescTreeView.Items.AddChild(DescTreeView.Items[nb2],Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Length(FinalList[nb])));
   end else if not HasBeenFound then begin
     // Sinon, on crée la branche et on y place le périphérique : [ici, le noeud parent est créé dans la fonction d'ajout de l'enfant]
     DevicesTreeView.Items.AddChild(  DevicesTreeView.Items.Add(nil, Copy(FinalList[nb],1,Pos('|',FinalList[nb])-1)) , Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Pos('§',FinalList[nb])-Pos('|',FinalList[nb])-1));
//                                   \_________________________création du noeud parent____________________________/
     // On crée le noeud équivalent dans le treeview d'infos
     DescTreeView.Items.AddChild(     DescTreeView.Items.Add(nil, Copy(FinalList[nb],1,Pos('|',FinalList[nb])-1))   ,Copy(FinalList[nb],Pos('|',FinalList[nb])+1,Length(FinalList[nb])));
//                                   \________________________création du noeud parent___________________________/
   end;

   // On évite le blocage de l'application
   Application.ProcessMessages;
  end;

 finally
  // On libère les éléments créés au départ
  Reg3.Free;
  Reg2.Free;
  Reg1.Free;
  FinalList.Free;
  SubSubCatList.Free;
  SubCatList.Free;
  CategoriesList.Free;
  // On réautorise l'énumération (voir ligne 94)
  EnumerateDevicesBtn.Enabled := true;
 end;
end;
//============================================================================//




//============================================================================//
// Attribution des images de catégories pour les noeuds racines et les enfants
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

 // On affiche le nom de la catégorie dans la partie informations si aucun périphérique particulier n'est sélectionné
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

 // Formatage préalable de la ligne (on découpe le début inutile)
 line := Copy(DescTreeView.Selected.Text,Pos('§', DescTreeView.Selected.Text),Length(DescTreeView.Selected.Text));

 // Et on remplit la StringGrid avec les informations
 i := 0;
 while Pos('§',line)<>0 do begin
  inc(i);
   PropertiesStringGrid.RowCount := i;
   PropertiesStringGrid.Cells[0,i-1] := Copy(line, 2, Pos('@', line)-2);
  if Pos('§',Copy(line,2,Length(line))) = 0 then begin
   PropertiesStringGrid.Cells[1,i-1] := Copy(line,Pos('@', line)+1,Length(line));
   line := '';
  end else begin
   PropertiesStringGrid.Cells[1,i-1] := Copy(line,Pos('@', line)+1,Pos('§',Copy(line,2,Length(line)))-Pos('@', line));
  end;
  // On supprime de la variable line l'information ajoutée
  line := Copy(line,Pos('§',Copy(line,2,Length(line)))+1,Length(line));
 end;
end;       

procedure TForm1.FormCreate(Sender: TObject);
begin
 PropertiesStringGrid.Cells[0,0] := 'device';
 PropertiesStringGrid.Cells[1,0] := 'properties';
end;

end.
