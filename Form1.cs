using System;
using System.Windows.Forms;

namespace for_AutoCAD_CSV_files
{
    public partial class AutoCADTextInserterForm : Form
    {
        // Свойство для хранения пути к выбранному файлу
        public string SelectedFilePath { get; private set; }

        public AutoCADTextInserterForm()
        {
            InitializeComponent();
        }

        private void button_choose_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            // Показать диалоговое окно открытия файла пользователю
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                // Если пользователь выбрал файл и нажал "OK", сохранить путь к файлу
                SelectedFilePath = openFileDialog.FileName;
                // Уведомление в консоль AutoCAD об успешном выборе файла
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nSelected file: {SelectedFilePath}");

                // Закрыть форму
                this.DialogResult = DialogResult.OK; // Устанавливаем результат диалога в OK
                this.Close();
            }
            else
            {
                // Если пользователь нажал "Отмена" или закрыл диалоговое окно
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nNo file selected.");
                this.DialogResult = DialogResult.Cancel; // Устанавливаем результат диалога в Cancel
            }
        }
    }
}