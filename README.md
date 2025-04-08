# Student Grade Calculator

This project calculates student grades based on scores from an Excel or text file. It provides an easy way to calculate and save grade results.

## How to Use

1. Click "Select File" to choose a file with student scores.
2. Click "Analyze" to calculate grades.
3. Results will be displayed in the text box.
4. You can save the results in a Word document.

## Features

- Supports both Excel and text files.
- Grades are calculated based on predefined criteria.
- Results can be saved as a Word document.

## Technologies Used

- PowerShell
- Excel Interop
- Word Interop

------------------------------------------------------------------------------------------------------------------------------------

Документация
Описание проекта
Этот скрипт предназначен для анализа данных студентов, включая их оценки. Программа позволяет пользователю выбрать файл с данными, проанализировать его и вывести результаты в окно. Также предоставляется возможность сохранить результаты в формате документа Word.

Основные компоненты
Форма Windows Forms:

Окно программы, в котором пользователь взаимодействует с кнопками и текстовыми полями.

Программа использует System.Windows.Forms для создания графического интерфейса.

Кнопки:

Select File: Открывает диалоговое окно для выбора файла.

Analyze: Начинает анализ данных из выбранного файла.

Clear Results: Очищает результаты.

Save as Word: Сохраняет результаты в формате документа Word.

Текстовые поля:

Поле для отображения пути к выбранному файлу.

Мультистрочное поле для отображения результатов анализа.

Обработка данных:

Программа поддерживает два типа файлов: текстовые и Excel.

Для текстовых файлов программа ожидает, что данные будут разделены пробелами, и последняя часть строки будет являться оценкой.

Для Excel файлов программа анализирует оценки студентов, расположенные в первом листе книги Excel.

Функции:

GetGrade: Определяет оценку по числовой отметке.

Как использовать
Запустите скрипт в PowerShell, убедившись, что у вас настроена политика выполнения скриптов (Set-ExecutionPolicy RemoteSigned).

Нажмите кнопку Select File, чтобы выбрать файл с данными студентов.

Нажмите кнопку Analyze, чтобы проанализировать файл и вывести результаты в текстовое поле.

Нажмите кнопку Save as Word, чтобы сохранить результаты в файл Word.

Если нужно очистить результаты, нажмите кнопку Clear Results.

Требования
Windows PowerShell 5.1 или выше.

.NET Framework для поддержки Windows Forms.

Microsoft Excel и Microsoft Word для работы с соответствующими файлами.

Примечания
В программе используется COM-объект для взаимодействия с Excel и Word. Убедитесь, что у вас установлены эти приложения.

Для корректной работы с файлами Excel и Word требуется установленный Office.
