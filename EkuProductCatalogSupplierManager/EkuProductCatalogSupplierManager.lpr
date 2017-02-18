Library EkuProductCatalogSupplierManager;

{$mode delphi}{$H+}

Uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms,
  ExtCtrls,
  EkuProductCatalogManager.Model.Manager,
  EkuProductCatalogSupplierManager.View.FormMain,
  EkuProductCatalogSupplierManager.View.FrameMain,
  Controls,
  SysUtils;

Var
  gModel: IManagerModel;

  Procedure PlugInInit(Parent: TWinControl; Model: IManagerModel); Stdcall;
  Begin
    Try
      Application.Initialize;
      FormMain := TFormMain.Create(Application);
      Try
        //FormMain.Parent := Parent;
        //FormMain.ParentWindow := Parent.Handle;
        //FormMain.Visible := True;
        //FormMain.Update();
        FormMain.ShowModal();
      Finally
        FormMain.Free();
      End;
    Except
      on E: Exception Do
      Begin
        //      WriteLn(E.Message);
      End;
    End;
  End;

Exports
  PlugInInit;

Begin
End.
