Unit EkuProductCatalogManager.Model.Manager;

{$mode delphi}{$H+}

Interface

Uses
  Classes, SysUtils;

Type

  { IManagerModel }

  IManagerModel = Interface(IInterface)
    ['{A6D0223D-90D4-490E-9D86-705EECF4F27E}']
    Procedure LoadPreference();
    Procedure SavePreference();
  End;


  { TManagerModel }

  TManagerModel = Class(TInterfacedObject, IManagerModel)
  Const
    CConfigFilePath: String = 'epcm.xml';
  Private
    FServerUri: String;
    FPassportUsername: String;
    FPassportPassword: String;
  Public
    Procedure LoadPreference();
    Procedure SavePreference();
    Property ServerUri: String Read FServerUri Write FServerUri;
    Property PassportUsername: String Read FPassportUsername Write FPassportUsername;
    Property PassportPassword: String Read FPassportPassword Write FPassportPassword;
  End;

Implementation

Uses
  laz2_DOM,
  laz2_XMLRead,
  laz2_XMLWrite,
  laz2_XMLUtils;

{ TManagerModel }

Procedure TManagerModel.LoadPreference;
Var
  mDoc: TXMLDocument;
Begin
  Try
    ReadXMLFile(mDoc, CConfigFilePath);
    FServerUri := mDoc.DocumentElement.FindNode('ServerUri').FirstChild.NodeValue;
    FPassportUsername := mDoc.DocumentElement.FindNode('PassportUsername').FirstChild.NodeValue;
    FPassportPassword := mDoc.DocumentElement.FindNode('PassportPassword').FirstChild.NodeValue;
  Finally
    FreeAndNil(mDoc);
  End;
End;

Procedure TManagerModel.SavePreference;
Var
  mDoc: TXMLDocument;
  mRootNode: TDOMNode;
  mTagNode: TDOMNode;
  mTextNode: TDOMNode;
Begin
  Try
    mDoc := TXMLDocument.Create();
    mRootNode := mDoc.CreateElement('Preference');
    mDoc.AppendChild(mRootNode);
    mRootNode := mDoc.DocumentElement;

    mTagNode := mDoc.CreateElement('ServerUri');
    mTextNode := mDoc.CreateTextNode(FServerUri);
    mTagNode.AppendChild(mTextNode);
    mRootNode.AppendChild(mTagNode);

    mTagNode := mDoc.CreateElement('PassportUsername');
    mTextNode := mDoc.CreateTextNode(FServerUri);
    mTagNode.AppendChild(mTextNode);
    mRootNode.AppendChild(mTagNode);

    mTagNode := mDoc.CreateElement('PassportPassword');
    mTextNode := mDoc.CreateTextNode(FServerUri);
    mTagNode.AppendChild(mTextNode);
    mRootNode.AppendChild(mTagNode);

    WriteXMLFile(mDoc, CConfigFilePath);
  Finally
    FreeAndNil(mDoc);
  End;
End;

End.

