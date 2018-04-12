#pragma once


namespace pop3_imap4_simplevc {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::Text;
	using namespace System::Text::RegularExpressions;
	using namespace EAGetMail;

	/// <summary>
	/// Summary for Form1
	///
	/// WARNING: If you change the name of this class, you will need to change the
	///          'Resource File Name' property for the managed resource compiler tool
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public ref class Form1 : public System::Windows::Forms::Form, public IComparer
	{
	public:
		Form1(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~Form1()
		{
			if (components)
			{
				delete components;
			}
		}
private: System::Windows::Forms::GroupBox ^groupBox1;
private: System::Windows::Forms::Label ^label1;
private: System::Windows::Forms::Label ^label2;
private: System::Windows::Forms::Label ^label3;
private: System::Windows::Forms::TextBox ^textServer;
private: System::Windows::Forms::TextBox ^textUser;
private: System::Windows::Forms::TextBox ^textPassword;
private: System::Windows::Forms::Label ^label4;
private: System::Windows::Forms::ComboBox ^lstAuthType;
private: System::Windows::Forms::Label ^label5;
private: System::Windows::Forms::ComboBox ^lstProtocol;
private: System::Windows::Forms::CheckBox ^chkLeaveCopy;
private: System::Windows::Forms::Button ^btnStart;
private: System::Windows::Forms::Button ^btnCancel;
private: System::Windows::Forms::Label ^lblStatus;
private: System::Windows::Forms::ProgressBar ^pgBar;
private: System::Windows::Forms::CheckBox ^chkSSL;
private: System::Windows::Forms::WebBrowser  ^webMail;
private: bool m_bcancel;
private: String ^m_uidlfile;
private: String ^m_curpath;
private: ArrayList ^m_arUidl;	
private: System::Windows::Forms::ListView ^lstMail;
private: System::Windows::Forms::ColumnHeader ^colFrom;
private: System::Windows::Forms::ColumnHeader ^colSubject;
private: System::Windows::Forms::ColumnHeader ^colDate;
private: System::Windows::Forms::Button ^btnDel;
private: System::Windows::Forms::Label ^lblTotal;
private: System::Windows::Forms::Label ^label6;

//EASendMail EventHandler
	private:
		System::Void OnConnected(Object ^sender, System::Boolean % cancel )
		{
			lblStatus->Text = "Connected ...";
			cancel = m_bcancel;
			Application::DoEvents();
		}

		System::Void OnQuit(Object ^sender, System::Boolean % cancel )
		{
			lblStatus->Text = "Quit ...";
			cancel = m_bcancel;
			Application::DoEvents();
		}

		System::Void OnReceivingDataStream(Object ^sender, MailInfo ^info, int received, int total, System::Boolean % cancel )
		{
			pgBar->Minimum =  0;
			pgBar->Maximum = total;
			pgBar->Value = received;
			cancel = m_bcancel;
			Application::DoEvents();
		}

		System::Void OnIdle( Object ^sender, System::Boolean % cancel )
		{
			cancel = m_bcancel;
			Application::DoEvents();
		}

		System::Void OnAuthorized( Object ^sender, System::Boolean % cancel )
		{
			lblStatus->Text =  "Authorized ...";
			cancel = m_bcancel;
			Application::DoEvents();
		}

		System::Void OnSecuring( Object ^sender, System::Boolean % cancel )
		{
			lblStatus->Text = "Securing ..." ;
			cancel = m_bcancel;
			Application::DoEvents();
		}

// Parse and Display Mails
		private: void LoadMails()
		{
			lstMail->Items->Clear();
			String ^mailFolder = String::Format( "{0}\\inbox", m_curpath );
			if( !Directory::Exists( mailFolder ))
				Directory::CreateDirectory( mailFolder );

			array<String^>^files  = Directory::GetFiles( mailFolder, "*.eml" );
			int count = files->Length;
			for( int i = 0; i < count; i++ )
			{
				String ^fullname = files[i];
				//For evaluation usage, please use "TryIt" as the license code, otherwise the 
				//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
				//"trial version expired" exception will be thrown.
				Mail ^oMail = gcnew Mail( "TryIt" );

				// Load( file, true ) only load the email header to Mail object to save the CPU and memory
				// the Mail object will load the whole email file later automatically if bodytext or attachment is required..
				oMail->Load( fullname, true );

				ListViewItem ^item = gcnew ListViewItem(oMail->From->ToString());
				item->SubItems->Add( oMail->Subject );
				item->SubItems->Add( oMail->ReceivedDate.ToString( "yyyy-MM-dd HH:mm:ss"));
				item->Tag = fullname;
				lstMail->Items->Add(item);

				int pos = fullname->LastIndexOf( ".");
				String ^mainName = fullname->Substring(0, pos);
				String ^htmlName = String::Format( "{0}.htm", mainName );
				if( !File::Exists( htmlName ))
				{
					// this email is unread, we set the font style to bold.
					item->Font = gcnew System::Drawing::Font(item->Font, FontStyle::Bold);
				}

				oMail->Clear();
			}
		}

		private: String^ _FormatHtmlTag( String ^src )
		{
			src = src->Replace( ">", "&gt;" );
			src = src->Replace( "<", "&lt;" );
			return src;
		}

		//we generate a html + attachment folder for every email, once the html is create,
		// next time we don't need to parse the email again.
		private: void _GenerateHtmlForEmail( String ^htmlName, String ^emlFile, String ^tempFolder )
		{
			//For evaluation usage, please use "TryIt" as the license code, otherwise the 
			//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
			//"trial version expired" exception will be thrown.
			Mail ^oMail = gcnew Mail("TryIt");
			oMail->Load( emlFile, false );

			if( oMail->IsEncrypted )
			{
				try
				{
					//this email is encrypted, we decrypt it by user default certificate.
					// you can also use specified certificate like this
					// Certificate *oCert = new Certificate();
					// oCert->Load(S"c:\\test.pfx", S"pfxpassword", Certificate::CertificateKeyLocation::CRYPT_USER_KEYSET)
					// oMail = oMail->Decrypt( oCert );
					oMail = oMail->Decrypt( nullptr );
				}
				catch(Exception ^ep )
				{
					MessageBox::Show( ep->Message );
					oMail->Load( emlFile, false );
				}	
			}

			if( oMail->IsSigned )
			{
				try
				{
					//this email is digital signed.
					EAGetMail::Certificate ^cert = oMail->VerifySignature();
					MessageBox::Show( "This email contains a valid digital signature.");
					//you can add the certificate to your certificate storage like this
					//cert->AddToStore( Certificate::CertificateStoreLocation::CERT_SYSTEM_STORE_CURRENT_USER,
					//	S"addressbook" );
					// then you can use send the encrypted email back to this sender.
				}
				catch(Exception ^ep )
				{
					MessageBox::Show( ep->Message );
				}
			}

			String ^html = oMail->HtmlBody;
			StringBuilder ^hdr = gcnew StringBuilder();

			hdr->Append( "<font face=\"Courier New,Arial\" size=2>" );
			hdr->Append( String::Format( "<b>From:</b> {0}<br>", _FormatHtmlTag(oMail->From->ToString())));
			array<MailAddress^>^  addrs = oMail->To;
			int count = addrs->Length;
			if( count > 0 )
			{
				hdr->Append( "<b>To:</b> ");
				for( int i = 0; i < count; i++ )
				{
					hdr->Append(  _FormatHtmlTag(addrs[i]->ToString()));
					if( i < count - 1 )
					{
						hdr->Append( ";" );
					}
				}
				hdr->Append( "<br>" );
			}

			addrs = oMail->Cc;

			count = addrs->Length;
			if( count > 0 )
			{
				hdr->Append( "<b>Cc:</b> ");
				for( int i = 0; i < count; i++ )
				{
					hdr->Append(  _FormatHtmlTag(addrs[i]->ToString()));
					if( i < count - 1 )
					{
						hdr->Append( ";" );
					}
				}
				hdr->Append( "<br>" );
			}

			hdr->Append( String::Format( "<b>Subject:</b>{0}<br>\r\n",  _FormatHtmlTag(oMail->Subject)));

			array<Attachment^>^ atts = oMail->Attachments;
			count = atts->Length;
			if( count > 0 )
			{
				if( !Directory::Exists( tempFolder ))
					Directory::CreateDirectory( tempFolder );

				hdr->Append( "<b>Attachments:</b>" );
				for( int i = 0; i < count; i++ )
				{
					Attachment ^att = atts[i];
					//this attachment is in OUTLOOK RTF format, decode it here.
					if( String::Compare( att->Name, "winmail.dat" ) == 0 )
					{
						array<Attachment^>^ tatts = nullptr;
						try
						{
							tatts = Mail::ParseTNEF( att->Content, true );
						}
						catch(Exception ^ep )
						{
							MessageBox::Show( ep->Message );
							continue;
						}

						int y = tatts->Length;
						for( int x = 0; x < y; x++ )
						{
							Attachment ^tatt = tatts[x];
							String ^tattname = String::Format( "{0}\\{1}", tempFolder, tatt->Name );
							tatt->SaveAs( tattname , true );
							hdr->Append( String::Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ", tattname, tatt->Name ));
						}
						continue;
					}

					String ^attname = String::Format( "{0}\\{1}", tempFolder, att->Name );
					att->SaveAs( attname , true );
					hdr->Append( String::Format( "<a href=\"{0}\" target=\"_blank\">{1}</a> ", attname, att->Name ));
					String^ simage = "image/";
					if( att->ContentID->Length > 0 )
					{	//show embedded image.
						html = html->Replace( String::Format( "cid:{0}", att->ContentID), attname );
					}
					else if( String::Compare( att->ContentType, 0, "image/", 0, simage->Length, true ) == 0 )
					{
						//show attached image.
						html = String::Concat( html, String::Format( "<hr><img src=\"{0}\">", attname ));
					}
				}
			}

			Regex ^reg = gcnew Regex( "(<meta[^>]*charset[ \t]*=[ \t\"]*)([^<> \r\n\"]*)",
				(RegexOptions)(RegexOptions::Multiline | RegexOptions::IgnoreCase));
			html = reg->Replace( html, "$1utf-8" );
			if( !reg->IsMatch( html ))
			{
				hdr->Insert( 0, "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text/html; charset=utf-8\">" );
			}

			html = html->Insert( 0, "<hr>" );
			html = html->Insert( 0, hdr->ToString());

			FileStream ^fs = gcnew FileStream(htmlName, FileMode::Create,FileAccess::Write, FileShare::None );
			array<unsigned char>^ data = System::Text::UTF8Encoding::UTF8->GetBytes( html );
			fs->Write(data, 0, data->Length);
			fs->Close();
			oMail->Clear();
		}

		private: void ShowMail( String ^fileName )
		{
			try
			{
				int pos = fileName->LastIndexOf(".");
				String ^mainName = fileName->Substring(0, pos);
				String ^htmlName = String::Format( "{0}.htm", mainName );
			
				String ^tempFolder = mainName;
				if(!File::Exists( htmlName ))
				{	//we haven't generate the html for this email, generate it now.
					_GenerateHtmlForEmail( htmlName, fileName, tempFolder );
				}

				Object ^empty = System::Reflection::Missing::Value;
				webMail->Navigate(htmlName);
			}
			catch( Exception ^ep )
			{
				MessageBox::Show( ep->Message );
			}
		}

	// sort the email as received data.
public: virtual int Compare(Object ^x, Object ^y)
		{
			ListViewItem ^itemx = dynamic_cast<ListViewItem^>(x);
			ListViewItem ^itemy = dynamic_cast<ListViewItem^>(y);

			String ^sx = itemx->SubItems[2]->Text;
			String ^sy = itemy->SubItems[2]->Text;

			if( sx->Length != sy->Length )
				return -1; //should never occured.

			int count = sx->Length;
			for( int i = 0; i < count; i++ )
			{
				if( sx[i] > sy[i] )
					return -1;
				else if( sx[i] < sy[i])
					return 1;
			}

			return 0;
		}
	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->groupBox1 = (gcnew System::Windows::Forms::GroupBox());
			this->pgBar = (gcnew System::Windows::Forms::ProgressBar());
			this->lblStatus = (gcnew System::Windows::Forms::Label());
			this->btnCancel = (gcnew System::Windows::Forms::Button());
			this->btnStart = (gcnew System::Windows::Forms::Button());
			this->chkLeaveCopy = (gcnew System::Windows::Forms::CheckBox());
			this->lstProtocol = (gcnew System::Windows::Forms::ComboBox());
			this->label5 = (gcnew System::Windows::Forms::Label());
			this->lstAuthType = (gcnew System::Windows::Forms::ComboBox());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->chkSSL = (gcnew System::Windows::Forms::CheckBox());
			this->textPassword = (gcnew System::Windows::Forms::TextBox());
			this->textUser = (gcnew System::Windows::Forms::TextBox());
			this->textServer = (gcnew System::Windows::Forms::TextBox());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->lstMail = (gcnew System::Windows::Forms::ListView());
			this->colFrom = (gcnew System::Windows::Forms::ColumnHeader());
			this->colSubject = (gcnew System::Windows::Forms::ColumnHeader());
			this->colDate = (gcnew System::Windows::Forms::ColumnHeader());
			this->btnDel = (gcnew System::Windows::Forms::Button());
			this->lblTotal = (gcnew System::Windows::Forms::Label());
			this->label6 = (gcnew System::Windows::Forms::Label());
			this->webMail = (gcnew System::Windows::Forms::WebBrowser());
			this->groupBox1->SuspendLayout();
			this->SuspendLayout();
			// 
			// groupBox1
			// 
			this->groupBox1->Controls->Add(this->pgBar);
			this->groupBox1->Controls->Add(this->lblStatus);
			this->groupBox1->Controls->Add(this->btnCancel);
			this->groupBox1->Controls->Add(this->btnStart);
			this->groupBox1->Controls->Add(this->chkLeaveCopy);
			this->groupBox1->Controls->Add(this->lstProtocol);
			this->groupBox1->Controls->Add(this->label5);
			this->groupBox1->Controls->Add(this->lstAuthType);
			this->groupBox1->Controls->Add(this->label4);
			this->groupBox1->Controls->Add(this->chkSSL);
			this->groupBox1->Controls->Add(this->textPassword);
			this->groupBox1->Controls->Add(this->textUser);
			this->groupBox1->Controls->Add(this->textServer);
			this->groupBox1->Controls->Add(this->label3);
			this->groupBox1->Controls->Add(this->label2);
			this->groupBox1->Controls->Add(this->label1);
			this->groupBox1->Location = System::Drawing::Point(8, 8);
			this->groupBox1->Name = L"groupBox1";
			this->groupBox1->Size = System::Drawing::Size(232, 328);
			this->groupBox1->TabIndex = 0;
			this->groupBox1->TabStop = false;
			this->groupBox1->Text = L"Account Information";
			// 
			// pgBar
			// 
			this->pgBar->Location = System::Drawing::Point(8, 312);
			this->pgBar->Name = L"pgBar";
			this->pgBar->Size = System::Drawing::Size(216, 8);
			this->pgBar->TabIndex = 15;
			// 
			// lblStatus
			// 
			this->lblStatus->AutoSize = true;
			this->lblStatus->Location = System::Drawing::Point(8, 272);
			this->lblStatus->Name = L"lblStatus";
			this->lblStatus->Size = System::Drawing::Size(0, 13);
			this->lblStatus->TabIndex = 14;
			// 
			// btnCancel
			// 
			this->btnCancel->Enabled = false;
			this->btnCancel->Location = System::Drawing::Point(128, 240);
			this->btnCancel->Name = L"btnCancel";
			this->btnCancel->Size = System::Drawing::Size(88, 24);
			this->btnCancel->TabIndex = 13;
			this->btnCancel->Text = L"Cancel";
			this->btnCancel->Click += gcnew System::EventHandler(this, &Form1::btnCancel_Click);
			// 
			// btnStart
			// 
			this->btnStart->Location = System::Drawing::Point(32, 240);
			this->btnStart->Name = L"btnStart";
			this->btnStart->Size = System::Drawing::Size(88, 24);
			this->btnStart->TabIndex = 12;
			this->btnStart->Text = L"Start";
			this->btnStart->Click += gcnew System::EventHandler(this, &Form1::btnStart_Click);
			// 
			// chkLeaveCopy
			// 
			this->chkLeaveCopy->Location = System::Drawing::Point(8, 208);
			this->chkLeaveCopy->Name = L"chkLeaveCopy";
			this->chkLeaveCopy->Size = System::Drawing::Size(208, 16);
			this->chkLeaveCopy->TabIndex = 11;
			this->chkLeaveCopy->Text = L"Leave a copy of message on server";
			// 
			// lstProtocol
			// 
			this->lstProtocol->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			this->lstProtocol->Location = System::Drawing::Point(80, 176);
			this->lstProtocol->Name = L"lstProtocol";
			this->lstProtocol->Size = System::Drawing::Size(136, 21);
			this->lstProtocol->TabIndex = 10;
			this->lstProtocol->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::lstProtocol_SelectedIndexChanged);
			// 
			// label5
			// 
			this->label5->AutoSize = true;
			this->label5->Location = System::Drawing::Point(8, 178);
			this->label5->Name = L"label5";
			this->label5->Size = System::Drawing::Size(46, 13);
			this->label5->TabIndex = 9;
			this->label5->Text = L"Protocol";
			// 
			// lstAuthType
			// 
			this->lstAuthType->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			this->lstAuthType->Location = System::Drawing::Point(80, 144);
			this->lstAuthType->Name = L"lstAuthType";
			this->lstAuthType->Size = System::Drawing::Size(136, 21);
			this->lstAuthType->TabIndex = 8;
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Location = System::Drawing::Point(8, 146);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(56, 13);
			this->label4->TabIndex = 7;
			this->label4->Text = L"Auth Type";
			// 
			// chkSSL
			// 
			this->chkSSL->Location = System::Drawing::Point(8, 120);
			this->chkSSL->Name = L"chkSSL";
			this->chkSSL->Size = System::Drawing::Size(208, 16);
			this->chkSSL->TabIndex = 6;
			this->chkSSL->Text = L"SSL connection";
			// 
			// textPassword
			// 
			this->textPassword->Location = System::Drawing::Point(80, 86);
			this->textPassword->Name = L"textPassword";
			this->textPassword->PasswordChar = '*';
			this->textPassword->Size = System::Drawing::Size(136, 20);
			this->textPassword->TabIndex = 5;
			// 
			// textUser
			// 
			this->textUser->Location = System::Drawing::Point(80, 54);
			this->textUser->Name = L"textUser";
			this->textUser->Size = System::Drawing::Size(136, 20);
			this->textUser->TabIndex = 4;
			// 
			// textServer
			// 
			this->textServer->Location = System::Drawing::Point(80, 22);
			this->textServer->Name = L"textServer";
			this->textServer->Size = System::Drawing::Size(136, 20);
			this->textServer->TabIndex = 3;
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Location = System::Drawing::Point(8, 88);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(53, 13);
			this->label3->TabIndex = 2;
			this->label3->Text = L"Password";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(8, 56);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(29, 13);
			this->label2->TabIndex = 1;
			this->label2->Text = L"User";
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(8, 24);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(38, 13);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Server";
			// 
			// lstMail
			// 
			this->lstMail->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(3) {this->colFrom, this->colSubject, 
				this->colDate});
			this->lstMail->FullRowSelect = true;
			this->lstMail->HideSelection = false;
			this->lstMail->Location = System::Drawing::Point(248, 16);
			this->lstMail->Name = L"lstMail";
			this->lstMail->Size = System::Drawing::Size(474, 168);
			this->lstMail->TabIndex = 1;
			this->lstMail->UseCompatibleStateImageBehavior = false;
			this->lstMail->View = System::Windows::Forms::View::Details;
			this->lstMail->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::lstMail_SelectedIndexChanged);
			// 
			// colFrom
			// 
			this->colFrom->Text = L"From";
			this->colFrom->Width = 100;
			// 
			// colSubject
			// 
			this->colSubject->Text = L"Subject";
			this->colSubject->Width = 200;
			// 
			// colDate
			// 
			this->colDate->Text = L"Date";
			this->colDate->Width = 150;
			// 
			// btnDel
			// 
			this->btnDel->Location = System::Drawing::Point(650, 186);
			this->btnDel->Name = L"btnDel";
			this->btnDel->Size = System::Drawing::Size(72, 24);
			this->btnDel->TabIndex = 3;
			this->btnDel->Text = L"Delete";
			this->btnDel->Click += gcnew System::EventHandler(this, &Form1::btnDel_Click);
			// 
			// lblTotal
			// 
			this->lblTotal->AutoSize = true;
			this->lblTotal->Location = System::Drawing::Point(256, 192);
			this->lblTotal->Name = L"lblTotal";
			this->lblTotal->Size = System::Drawing::Size(0, 13);
			this->lblTotal->TabIndex = 4;
			// 
			// label6
			// 
			this->label6->ForeColor = System::Drawing::SystemColors::HotTrack;
			this->label6->Location = System::Drawing::Point(16, 352);
			this->label6->Name = L"label6";
			this->label6->Size = System::Drawing::Size(216, 48);
			this->label6->TabIndex = 5;
			this->label6->Text = L"Warning: if \"leave a copy of message on server\" is not checked,  the emails on th" 
				L"e server will be deleted !";
			// 
			// webMail
			// 
			this->webMail->Location = System::Drawing::Point(248, 216);
			this->webMail->Name = L"webMail";
			this->webMail->Size = System::Drawing::Size(474, 176);
			this->webMail->TabIndex = 1;
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(734, 412);
			this->Controls->Add(this->label6);
			this->Controls->Add(this->lblTotal);
			this->Controls->Add(this->btnDel);
			this->Controls->Add(this->webMail);
			this->Controls->Add(this->lstMail);
			this->Controls->Add(this->groupBox1);
			this->Name = L"Form1";
			this->Text = L"Form1";
			this->Load += gcnew System::EventHandler(this, &Form1::Form1_Load);
			this->Resize += gcnew System::EventHandler(this, &Form1::Form1_Resize);
			this->groupBox1->ResumeLayout(false);
			this->groupBox1->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void Form1_Load(System::Object^  sender, System::EventArgs^  e) {

				 m_bcancel = false;
				 m_uidlfile = "uidl.txt";
				 m_curpath = "";
				 m_arUidl = gcnew ArrayList();
				 
				 webMail->Navigate( "about:blank" ); 

				 lstProtocol->Items->Add( "POP3" );
				 lstProtocol->Items->Add( "IMAP4" );
				 lstProtocol->Items->Add("Exchange Web Service - 2007/2010");
				 lstProtocol->Items->Add("Exchange WebDAV - Exchange 2000/2003");
				 
				 lstProtocol->SelectedIndex = 0;

				 lstAuthType->Items->Add( "USER/LOGIN" );
				 lstAuthType->Items->Add( "APOP" );
				 lstAuthType->Items->Add( "NTLM" );
				 lstAuthType->SelectedIndex = 0;

				 String ^path = Application::ExecutablePath;
				 int pos = path->LastIndexOf( '\\' );
				 if( pos != -1 )
					 path = path->Substring( 0, pos );

				 m_curpath = path;

				 lstMail->Sorting = SortOrder::Descending;
				 lstMail->ListViewItemSorter = this;

				 LoadMails();
				 lblTotal->Text = String::Format( "Total {0} email(s)", lstMail->Items->Count );
			 }
private: System::Void btnCancel_Click(System::Object^  sender, System::EventArgs^  e) {
			 m_bcancel = true;
		 }
private: System::Void btnDel_Click(System::Object^  sender, System::EventArgs^  e) {
			
			 ListView::SelectedListViewItemCollection ^items = lstMail->SelectedItems;
			if( items->Count == 0 )
				return;

			if( MessageBox::Show("Do you want to delete all selected emails", 
                             "", 
							 MessageBoxButtons::YesNo) == Windows::Forms::DialogResult::No )
				 return;

			while( items->Count > 0 )
			{
				try
				{
					String ^fileName = dynamic_cast<String^>(items[0]->Tag);
					File::Delete( fileName );
					int pos = fileName->LastIndexOf( ".");
					String ^tempFolder = fileName->Substring(0, pos);
					String ^htmlName = String::Format( "{0}.htm",  tempFolder );
					if( File::Exists( htmlName ))
						File::Delete( htmlName );
					
					if( Directory::Exists( tempFolder ))
					{
						Directory::Delete( tempFolder, true );
					}

					lstMail->Items->Remove(items[0]);
				}
				catch(Exception ^ep )
				{
					MessageBox::Show( ep->Message );
					break;
				}
			}

			lblTotal->Text = String::Format( "Total {0} email(s)", lstMail->Items->Count);
			webMail->Navigate( "about:blank");
		 }
private: System::Void btnStart_Click(System::Object^  sender, System::EventArgs^  e) {
			String ^server, ^user, ^password;
			server = textServer->Text->Trim();
			user = textUser->Text->Trim();
			password = textPassword->Text->Trim();

			if( server->Length == 0 || user->Length == 0 || password->Length == 0 )
			{
				MessageBox::Show("Please input server, user and password." );
				return;
			}

			btnStart->Enabled = false;
			btnCancel->Enabled = true;

			ServerAuthType authType = ServerAuthType::AuthLogin;
			if( lstAuthType->SelectedIndex == 1 )
				authType = ServerAuthType::AuthCRAM5;
			else if( lstAuthType->SelectedIndex == 2 )
				authType = ServerAuthType::AuthNTLM;

			ServerProtocol protocol = (ServerProtocol)lstProtocol->SelectedIndex;

			MailServer ^oServer = gcnew MailServer( server, user, password,
				chkSSL->Checked, authType, protocol );

            //For evaluation usage, please use "TryIt" as the license code, otherwise the 
            //"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            //"trial version expired" exception will be thrown.
			MailClient ^oClient = gcnew MailClient("TryIt");

			//Catching the following events is not necessary, 
            //just make the application more user friendly.
            //If you use the object in asp.net/windows service or non-gui application, 
            //You need not to catch the following events.
            //To learn more detail, please refer to the code in EAGetMail EventHandler region
			oClient->OnAuthorized += gcnew MailClient::OnAuthorizedEventHandler( this,  &Form1::OnAuthorized );
			oClient->OnConnected += gcnew MailClient::OnConnectedEventHandler( this,  &Form1::OnConnected );
			oClient->OnIdle += gcnew MailClient::OnIdleEventHandler( this,  &Form1::OnIdle );
			oClient->OnSecuring += gcnew MailClient::OnSecuringEventHandler( this,  &Form1::OnSecuring );
			oClient->OnReceivingDataStream += gcnew MailClient::OnReceivingDataStreamEventHandler( this,  &Form1::OnReceivingDataStream );	

			bool bLeaveCopy = chkLeaveCopy->Checked;
			
			// UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
            // the same email from server more than once, we record the email UIDL retrieved every time
            // if you delete the email from server every time and not to leave a copy of email on
            // the server, then please remove all the function about uidl.
            // UIDLManager wraps the function to write/read uidl record from a text file.
            UIDLManager^ oUIDLManager  = gcnew UIDLManager();

			try
			{
				// load existed uidl records to UIDLManager
				String^ uidlfile = String::Format("{0}\\{1}", m_curpath, m_uidlfile);
                oUIDLManager->Load(uidlfile);


				String ^mailFolder = String::Format( "{0}\\inbox", m_curpath );
				if( !Directory::Exists( mailFolder ))
					Directory::CreateDirectory( mailFolder );

				m_bcancel = false;
				lblStatus->Text = "Connecting ...";
				oClient->Connect( oServer );
				array<MailInfo^>^infos = oClient->GetMailInfos();
				lblStatus->Text = String::Format( "Total {0} email(s)" ,infos->Length);
				
				// remove the local uidl which is not existed on the server.
                oUIDLManager->SyncUIDL(oServer, infos);
                oUIDLManager->Update();

				int count = infos->Length;

				for( int i = 0; i < count; i++ )
				{
					MailInfo ^info = infos[i];
					if(oUIDLManager->FindUIDL( oServer, info->UIDL ) != nullptr )
					{
						// This email has been downloaded on local disk.
						continue;
					}

					lblStatus->Text = String::Format( "Retrieving {0}/{1}...",info->Index, count);
					
					Mail ^oMail = oClient->GetMail( info );
					System::DateTime d = System::DateTime::Now;
					
					System::Globalization::CultureInfo ^cur = gcnew System::Globalization::CultureInfo("en-US");			
					String ^sdate = d.ToString("yyyyMMddHHmmss", cur);
					String ^fileName = String::Format( "{0}\\{1}{2}", 
						mailFolder, 
						sdate, 
						d.Millisecond.ToString("d3"));

					fileName = String::Format( "{0}{1}.eml", fileName, i);
						
					oMail->SaveAs( fileName, true );

					ListViewItem ^item = gcnew ListViewItem(oMail->From->ToString());
					item->SubItems->Add( oMail->Subject );
					item->SubItems->Add( oMail->ReceivedDate.ToString( "yyyy-MM-dd HH:mm:ss"));
					item->Font = gcnew System::Drawing::Font(item->Font, FontStyle::Bold);
					item->Tag = fileName;
					lstMail->Items->Insert(0, item);
					oMail->Clear();

					lblTotal->Text = String::Format( "Total {0} email(s)", lstMail->Items->Count);

					if( bLeaveCopy )
					{
						// Add the email uidl to uidl file to avoid we retrieve it next time. 
						oUIDLManager->AddUIDL( oServer, info->UIDL, fileName );
					}
				}

				if( !bLeaveCopy )
				{
					lblStatus->Text = "Deleting ...";
					for( int i = 0; i < count; i++ )
					{
						oClient->Delete( infos[i] );
						// Remove UIDL from local uidl file.
                        oUIDLManager->RemoveUIDL( oServer, infos[i]->UIDL );
					}
				}
				// Delete method just mark the email as deleted, 
				// Quit method pure the emails from server exactly.
				oClient->Quit();
			
			}
			catch( Exception ^ep )
			{
				MessageBox::Show( ep->Message );
			}
				
			// Update the uidl list to local uidl file and then we can load it next time.
			oUIDLManager->Update();

			lblStatus->Text = "Completed";
			pgBar->Maximum = 100;
			pgBar->Minimum = 0;
			pgBar->Value = 0;
			btnStart->Enabled = true;
			btnCancel->Enabled = false;
		 }
private: System::Void lstMail_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {

			ListView::SelectedListViewItemCollection ^items = lstMail->SelectedItems;
			if( items->Count == 0 )
				return;

			ListViewItem ^item = items[0];
			ShowMail( dynamic_cast<String^>(item->Tag));
			item->Font = gcnew System::Drawing::Font(item->Font, FontStyle::Regular);

		 }
private: System::Void Form1_Resize(System::Object^  sender, System::EventArgs^  e) {

			if (this->Width < 750)
            {
                this->Width = 750;
            }

            if (this->Height < 450)
            {
                this->Height = 450;
            }

            lstMail->Width = this->Width - 270;
            webMail->Width = lstMail->Width;
            btnDel->Left = this->Width - (btnDel->Width + 20);
            webMail->Height = this->Height - (lstMail->Height + 100);
		 }
private: System::Void lstProtocol_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {

			// By default, Exchange Web Service requires SSL connection.
            // For other protocol, please set SSL connection based on your server setting manually
            {
				if (lstProtocol->SelectedIndex == (int)ServerProtocol::ExchangeEWS)
                {
                    chkSSL->Checked = true;
                }
            }
		 }
};
}

