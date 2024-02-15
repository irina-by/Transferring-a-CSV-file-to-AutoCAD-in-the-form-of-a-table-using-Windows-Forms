using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.IO;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace for_AutoCAD_CSV_files
{
    public class CSVReaderCommands
    {
        [CommandMethod("OpenCSV")]
        public void OpenCSVCommand()
        {
            // Получение ссылок на активный документ и редактор из AutoCAD.
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Создание и показ формы AutoCADTextInserterForm
            AutoCADTextInserterForm form = new AutoCADTextInserterForm();

            // Показываем форму как модальное диалоговое окно и сохраняем результат в переменной diaRes
            DialogResult diaRes = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form);

            // Проверяем результат
            if (diaRes == DialogResult.OK)
            {
                //код для чтения файла и получения данных
                // Если пользователь нажал OK, получаем путь к выбранному файлу
                string filePath = form.SelectedFilePath; //хранит путь к CSV-файлу, который был получен от пользователя
                ed.WriteMessage($"\nПуть к выбранному файлу: {filePath}");
                // Чтение CSV файла по указанному пути.
                string[][] csvData = ReadCSVFile(filePath);  // это двумерный массив строк, который получается в результате чтения CSV-файла по пути filePath,каждый элемент массива представляет собой строку из CSV файла.
                                                             //Получение заголовочной строки CSV файла. 
                string[] headerRow = csvData[0]; //Заголовочная строка представляет собой первую строку в csvData массиве. Она извлекается и сохраняется в массиве headerRow.
                                                 //Получение координат для размещения текста в AutoCAD.
                double[] coordinates = GetTableCorner(); //содержит координаты, которые были получены от пользователя, для определения начальной точки размещения текста в автокад
                Point3d[][] points = GetPositions(csvData, coordinates[0], coordinates[1]);//это двумерный массив объектов Point3d, который хранит точки, где будет размещен текст
                Point3d[] headerCoordinates = points[0];

                
                // Создание списка слов List<string> words для хранения данных из CSV
                List<string> words = new List<string>(); //это лист, который будет хранить все ячейки данных csvData (это двумерный массив строк, который был считан из CSV-файла)в одном линейном списке 

                
                for (int i = 1; i < csvData.Length; i++) //Цикл начинается с i = 1, что предполагает, что первая строка (i = 0) содержит заголовки столбцов и не должна включаться в список слов
                {
                    for (int j = 0; j < csvData[i].Length; j++) //Вложенные циклы проходят по всем строкам и столбцам массива csvData (за исключением первой строки)
                    {
                        words.Add(csvData[i][j]); // все элементы данных добавляются в список words 
                    }
                }

                //Вывод текста в документ AutoCAD
                int k = 0;
                for (int i = 0; i < points.Length; i++) //points - это двумерный массив объектов Point3d, который содержит координаты для размещения каждой ячейки текста в AutoCAD
                {
                    for (int j = 0; j < points[i].Length; j++)
                    {
                        if (k < words.Count) // Проверка, что индекс не выходит за пределы списка слов
                        {
                            if (points[i][j].Y == coordinates[1]) // проверяет, находится ли текущая точка в той же Y-координате, что и левый верхний угол таблицы (то есть это первая строка для заголовков).
                            {
                                DrawMTextHeader(headerRow[j], points[i][j]);//headerRow[j] - это массив с текстом заголовков столбцов из первой строки csvData.Если условие истинно, вызывается метод DrawMTextHeader, который рисует заголовок столбца в этой позиции
                            }
                            else //Если условие ложно (то есть для всех остальных строк таблицы), вызывается метод DrawMText для отображения данных из списка words
                            {
                                DrawMText(words[k], points[i][j], false);//После каждого вызова DrawMText, индекс k увеличивается, чтобы перейти к следующему слову в списке words
                                k++; // Переходим к следующему слову в списке
                            }
                        }
                    }
                }
            }
            else
            {
                // Если пользователь не нажал OK, выходим из команды
                ed.WriteMessage("\nОткрытие CSV-файла отменено пользователем.");
            }
        }

        

        // Метод для чтения данных из CSV файла
        private string[][] ReadCSVFile(string filePath)
        {
            List<string[]> temp = new List<string[]>();
            string[][] csvData = new string[temp.Count][];
            try
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] row = line.Split(new string[] { ";" }, StringSplitOptions.None);
                        temp.Add(row);
                    }
                }
                csvData = new string[temp.Count][];
                for (int i = 0; i < csvData.Length; i++)
                {
                    csvData[i] = temp[i];

                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("Error reading from CSV file: " + e.Message);
            }

            return csvData;
        }

        //Получение с консоли количества строк в таблице

        [CommandMethod("NumberOfLines")]
        public int PromptForNumberOfRows()
        {
            // Получаем ссылку на текущий документ и его редактор
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor Ed = acDoc.Editor;

            // Запрашиваем у пользователя ввести количество строк
            PromptIntegerOptions promptOptions = new PromptIntegerOptions("\nВведите количество строк в таблице: ")
            {
                AllowNegative = false, // Не разрешаем отрицательные значения
                AllowZero = false // Не разрешаем ноль
            };

            // Получаем результат ввода пользователя
            PromptIntegerResult promptResult = Ed.GetInteger(promptOptions);
            int NumberOfRowsInTheTable = 1;
            // Проверяем результат ввода
            if (promptResult.Status == PromptStatus.OK)
            {
                // Если пользователь ввел число, присваиваем его NumberOfRowsInTheTable
                NumberOfRowsInTheTable = promptResult.Value;
            }
            else
            {
                // Если пользователь отменил ввод, выводим сообщение об отмене
                Ed.WriteMessage("\nВвод количества строк отменен.");
            }
            return NumberOfRowsInTheTable;
        }

        private Point3d[][] GetPositions(string[][] csvData, double x, double y)
        {
            int[] coordX = { 0, 20, 45, 80, 120, 162, 192, 207, 230, 275, 322, 337, 352, 367, 382, 400, 417, 430, 440, 452, 467, 482, 495, 505, 517, 532, 547, 565 };
            int coordY = 0;
            int shift_x = 0;
            int checkNewTable = shift_x;

            int numberofrows = PromptForNumberOfRows(); //назначаем до цикла фор чтобы постоянно не выводился метод
            int distancerequest = GetSpacingBetweenTables();
            double temp = csvData.Length / numberofrows;
            Point3d[][] coordinates = new Point3d[csvData.Length + Convert.ToInt32(Math.Ceiling(temp))][];//Создание двумерного массива coordinates, размер которого равен длине входного массива csvData.
            for (int i = 0; i < coordinates.Length; i++) //Внешний цикл for перебирает каждую строку в csvData
            {
                coordinates[i] = new Point3d[csvData[0].Length];//Для каждой строки создается массив Point3d, размер которого равен количеству элементов в этой строке csvData[i]

                shift_x = distancerequest * (i / numberofrows); //600 сдвиг в другую таблицу, numberofrows-число из файла настроек
                if (checkNewTable != shift_x)
                {
                    coordY = 0;
                }
                for (int j = 0; j < coordinates[i].Length; j++)//Внутренний цикл for перебирает каждый элемент текущей строки csvData[i].
                {
                    Point3d point = new Point3d(coordX[j] + x + shift_x, coordY + y, 0); //Внутри внутреннего цикла создается новый объект Point3d с координатами X, Y и Z,Координата X объекта Point3d вычисляется путем добавления смещения x к значению из массива coordX[j].
                                                                                         //относительно края таблицы, Координата Y объекта Point3d равна текущему значению coordY плюс смещение y- начальная точка при объявлении GetTableCorner.
                                                                                         //Координата Z объекта Point3d устанавливается равной 0, что указывает, что точки лежат в плоскости XY.
                    coordinates[i][j] = point; //После создания, объект Point3d добавляется в текущий массив точек coordinates[i]
                }
                if (checkNewTable != shift_x || i == 0) // сдвиг от заголовка
                {
                    coordY = coordY - 30;
                }
                else
                {
                    coordY = coordY - 22; //сдвигает следующий ряд точек на 22 единиц вниз по оси Y
                }

                checkNewTable = shift_x;
            }
            return coordinates;
        }

        //Получение точки для вставки таблицы
        [CommandMethod("GetTableCorner")]
        public double[] GetTableCorner()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;

            // Запрос координат у пользователя
            PromptPointResult pPointRes;
            PromptPointOptions pPointOpts = new PromptPointOptions("");

            // Установка сообщения запроса
            pPointOpts.Message = "\nУкажите точку для вставки таблицы: ";

            // Получение координат от пользователя
            pPointRes = ed.GetPoint(pPointOpts);
            Point3d tableCorner;
            double[] coordinates = { 0, 0 };
            // Проверка на отмену запроса пользователем
            if (pPointRes.Status == PromptStatus.OK)
            {
                tableCorner = pPointRes.Value;

                // Извлекаем координаты X, Y и Z
                double x = tableCorner.X;
                double y = tableCorner.Y;
                double z = tableCorner.Z;
                coordinates = new double[] { x, y };

                // Теперь координаты можно использовать для дальнейших операций
                ed.WriteMessage($"Координаты для вставки таблицы: X = {x}, Y = {y}, Z = {z}\n");
            }
            else
            {
                ed.WriteMessage("Команда отменена пользователем.\n");
            }
            return coordinates;

        }

        // Метод для запроса расстояния между таблицами
        public int GetSpacingBetweenTables()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;

            PromptDistanceOptions pdo = new PromptDistanceOptions("\nУкажите расстояние между таблицами: ")
            {
                AllowNegative = false, // Не разрешаем отрицательные значения
                AllowZero = false // Не разрешаем ноль
            };

            PromptDoubleResult pdr = ed.GetDistance(pdo);

            if (pdr.Status == PromptStatus.OK)
            {
                // Преобразуем и возвращаем значение расстояния как целое число
                return (int)pdr.Value;
            }
            else
            {
                // Если пользователь отменил ввод, выводим сообщение об отмене и возвращаем -1
                ed.WriteMessage("\nОтмена ввода расстояния.");
                return -1;
            }
        }

        //Метод использует транзакции для безопасного создания и вставки объекта многострочного текста (MText) в документ AutoCAD
        public void DrawMText(string text, Point3d point, bool isRotate)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor edt = doc.Editor;

            using (Transaction trans = db.TransactionManager.StartTransaction()) //создает новую транзакцию, используя менеджер транзакций базы данных (db)
            {
                try
                {
                    BlockTable bt;
                    bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable; //получаем доступ к таблице блоков (BlockTable)

                    BlockTableRecord btr;
                    btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;//Это обращение к записи таблицы блоков для модельного пространства (ModelSpace), которая содержит все графические объекты. Запись открывается в режиме ForWrite для записи, так как мы собираемся вносить изменения.

                    string txt = text; //переменная, содержащая текст
                    Point3d insPt = point; //это точка с координатами X, Y и Z
                    using (MText mtx = new MText()) //Создание объекта MText для отображения нашего текста
                    {
                        mtx.Contents = txt;
                        mtx.Location = insPt;
                        SetSPDSTextStyle(mtx, db, trans, insPt, isRotate); //стиль текста для объекта MText

                        btr.AppendEntity(mtx); //Объект MText добавляется в модельное пространство
                        trans.AddNewlyCreatedDBObject(mtx, true);//объект был создан и его надо включить в базу данных
                    }
                    trans.Commit();//В случае успешного выполнения операций, изменения фиксируются в базе данных с помощью метода Commit. 

                }
                catch (System.Exception ex)
                {
                    edt.WriteMessage("Error!" + ex.Message);
                    trans.Abort();
                }
            }
        }


        //стиль текста
        private void SetSPDSTextStyle(MText text, Database db, Transaction tr, Point3d position, bool isRotate)
        {
            TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

            ObjectId styleId = tst["SPDS"];

            text.Height = 2.5;
            text.TextStyleId = styleId;

            text.Width = 37;//ширина
            text.Height = 0;//высота

            // Установка выравнивания текста вверх по центру
            text.Attachment = AttachmentPoint.TopCenter;

            //Поворот текста на 90 градусов
            if (isRotate)
            {
                text.Rotation = Math.PI / 2.0;
            }

        }
        
        
        //заголовок
        private void DrawMTextHeader(string text, Point3d point)
        {
            DrawMText(text, point, true);

        }


        // Метод для фокусировки на объекте MText
        private void ZoomToMText(MText mText)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Устанавливаем границы для видового окна
            Extents3d extents = mText.GeometricExtents;
            Point3d minPoint = extents.MinPoint;
            Point3d maxPoint = extents.MaxPoint;

            // Вычисляем центр
            Point3d midPoint = minPoint + ((maxPoint - minPoint) * 0.5);

            // Вычисляем размеры "окна" просмотра
            double width = (maxPoint.X - minPoint.X);
            double height = (maxPoint.Y - minPoint.Y);

            // Создаем новое видовое окно, центрированное на мультитексте
            ViewTableRecord view = new ViewTableRecord();
            view.CenterPoint = new Point2d(midPoint.X, midPoint.Y);
            view.Height = height;
            view.Width = width;

            // Применяем новое видовое окно
            ed.SetCurrentView(view);
        }

    }
}
    

