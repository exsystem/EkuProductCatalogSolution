Unit EkuProductCatalogManager.Controller.Manager;

{$mode delphi}{$H+}

Interface

Uses
  Classes, SysUtils, Forms, Dialogs, ExtCtrls, Controls,
  EkuProductCatalogManager.Model.Manager,
  EkuProductCatalogManager.View.FormMain;

Type

  { TManagerController }

  TManagerController = Class(TInterfacedObject)
  Const
    CPlugInInitProcName = 'PlugInInit';
  Private
    FApplication: TApplication;
    FModel: IManagerModel;
    FView: TFormMain;
    FPlugIn: TLibHandle;
  Public
    Constructor Create(App: TApplication; Model: TManagerModel; View: TFormMain);
    Destructor Destroy; Override;
    Procedure LoadPlugIn(Name: String);
  End;

  TPlugInInitProc = Procedure(Parent: TWinControl; Model: IManagerModel); Stdcall;

Var
  Controller: TManagerController;

Implementation

{ TManagerController }

Constructor TManagerController.Create(App: TApplication; Model: TManagerModel; View: TFormMain);
Begin
  FApplication := Application;
  FModel := Model;
  FView := View;
End;

Destructor TManagerController.Destroy;
Begin
  UnloadLibrary(FPlugIn);
  FModel := Nil;
  Inherited Destroy;
End;

Procedure TManagerController.LoadPlugIn(Name: String);
Var
  mProc: TPlugInInitProc;
Begin
  If (FPlugIn <> NilHandle) Then
  Begin
    UnloadLibrary(FPlugIn);
  End;
  FPlugIn := LoadLibrary(Name);
  @mProc := GetProcedureAddress(FPlugIn, CPlugInInitProcName);
  If @mProc = Nil Then
  Begin
    MessageDlg('錯誤', '載入插件錯誤！', mtError, [mbOK], '');
    Exit();
  End;

  Try
    mProc(FView.PanelMain, FModel);
  Except
    On E: Exception Do
      MessageDlg('ERROR', E.Message, mtError, [mbOK], '');
  End;
End;

End.

