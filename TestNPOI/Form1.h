#pragma once


namespace TestNPOI {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;

	// Los using del NPOI
	using namespace NPOI::SS::UserModel;
	using namespace NPOI::HSSF::Model;
	using namespace NPOI::HSSF::UserModel;

	/// <summary>
	/// Summary for Form1
	///
	/// WARNING: If you change the name of this class, you will need to change the
	///          'Resource File Name' property for the managed resource compiler tool
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public ref class Form1 : public System::Windows::Forms::Form
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
	private: System::Windows::Forms::Button^  button1;
	protected: 

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
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(34, 59);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(75, 23);
			this->button1->TabIndex = 0;
			this->button1->Text = L"button1";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &Form1::button1_Click);
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(599, 366);
			this->Controls->Add(this->button1);
			this->Name = L"Form1";
			this->Text = L"Form1";
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
				 //	Quiero escribir una celda
				 if(!File::Exists(Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments) + "\\prueba.xls")){
					 HSSFWorkbook^ libro = HSSFWorkbook::Create(InternalWorkbook::CreateWorkbook());
					 HSSFSheet ^hoja = dynamic_cast<HSSFSheet^>(libro->CreateSheet("Hoja1"));
					 for(int i = 0; i < 10; i++){
						 IRow^ fila = hoja->CreateRow(i);
						 for(int j = 0; j < 10; j++){
							 ICell ^c = fila->CreateCell(j);
							 c->SetCellValue("Esto es una prueba de npoi:" + i.ToString() + " , " + j.ToString() );
						 }
					 }

					 try{
						 FileStream^ fs = gcnew FileStream(Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments) + "\\prueba.xls", FileMode::Create, FileAccess::Write);
						 libro->Write(fs);
						 fs->Close();
					 }catch(Exception^ ex){
						 MessageBox::Show("Error al crear el archivo: " + ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
					 }
				 }
			 }
	};
}

