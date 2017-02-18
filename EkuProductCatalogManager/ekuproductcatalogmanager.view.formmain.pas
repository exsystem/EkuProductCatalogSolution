Unit EkuProductCatalogManager.View.FormMain;

{$mode objfpc}{$H+}

Interface

Uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Menus,
  ExtCtrls,
  EkuProductCatalogManager.Model.Manager;

Type

  { TFormMain }

  TFormMain = Class(TForm)
    MainMenuMain: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItemTool: TMenuItem;
    MenuItemToolPreference: TMenuItem;
    PanelMain: TPanel;
    Procedure MenuItem1Click(Sender: TObject);
    Procedure MenuItem2Click(Sender: TObject);
  Private

  Public

  End;

Procedure t(Parent: TWinControl; Model: IManagerModel); Stdcall; External '..\EkuProductCatalogSupplierManager\EkuProductCatalogSupplierManager.dll' Name 'PlugInInit'; //(***) COMMENT HERE

Var
  FormMain: TFormMain;

Implementation

Uses
  EkuProductCatalogManager.Controller.Manager;

{$R *.lfm}

{ TFormMain }

Procedure TFormMain.MenuItem1Click(Sender: TObject);
Begin
  Controller.LoadPlugIn('..\EkuProductCatalogSupplierManager\EkuProductCatalogSupplierManager.dll');
End;

Procedure TFormMain.MenuItem2Click(Sender: TObject);
Begin
  t(PanelMain, Nil); //(***) COMMENT HERE
End;

End.
