Program EkuProductCatalogManager;

{$mode delphi}{$H+}

Uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms,
  EkuProductCatalogManager.View.FormMain,
  EkuProductCatalogManager.Model.Manager,
  EkuProductCatalogManager.Controller.Manager { you can add units after this };

{$R *.res}

Begin
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Controller := TManagerController.Create(Application, TManagerModel.Create(), FormMain);
  Try
    Application.Run;
  Finally
    Controller.Free();
  End;
End.

