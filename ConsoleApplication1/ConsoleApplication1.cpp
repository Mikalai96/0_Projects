
#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <filesystem>
#include <ctime>
#include <iomanip>
#include "correspondence.h"
#include "journal.h"
#include "file_manager.h"
#include "config_manager.h"
#include <Windows.h>
#include "pdfManager.h"

namespace fs = std::filesystem;

void displayMenu() {
    std::cout << "\n===== Файловый менеджер корреспонденции =====\n";
    std::cout << "1. Зарегистрировать входящий документ\n";
    std::cout << "2. Зарегистрировать исходящий документ\n";
    std::cout << "3. Просмотреть журнал корреспонденции\n";
    std::cout << "4. Поиск документа\n";
    std::cout << "5. Изменить место хранения данных\n";
    std::cout << "6. Экспорт журнала в TXT\n";
    std::cout << "7. Открытие прикрепленного файла\n";
    std::cout << "0. Выход\n";
    std::cout << "Выберите действие: ";
}

std::string getCurrentDate() {
    auto now = std::time(nullptr);
    std::tm timeinfo;
    localtime_s(&timeinfo, &now);
    std::ostringstream oss;
    oss << std::put_time(&timeinfo, "%d.%m.%Y");
    return oss.str();
}

void promptUserForStoragePath(FileManager& manager, ConfigManager& configManager) {
    std::string userPath;
    std::cout << "Введите путь для хранения информации о корреспонденции: ";
    std::getline(std::cin, userPath);

    try {
        manager.setStoragePath(userPath);
        std::cout << "Путь успешно установлен: " << manager.getStoragePath() << std::endl;
        configManager.saveStoragePath(manager.getStoragePath());
    }
    catch (const std::exception& e) {
        std::cerr << "Ошибка: " << e.what() << std::endl;
        std::cout << "Используется путь по умолчанию." << std::endl;
    }
}



int main() {
    SetConsoleCP(1251);       // Установка кодировки ввода
    SetConsoleOutputCP(1251); // Установка кодировки вывода                            
    setlocale(LC_ALL, "Russian"); // Установка русской локали

    // Инициализация менеджеров
    FileManager fileManager;
    ConfigManager configManager;

    // Загрузка сохраненного пути
    std::string savedPath = configManager.loadStoragePath();
    if (!savedPath.empty()) {
        try {
            fileManager.setStoragePath(savedPath);
            std::cout << "Загружен сохраненный путь: " << savedPath << std::endl;
        }
        catch (const std::exception& e) {
            std::cout << "Ошибка при использовании сохраненного пути. Запрос нового пути." << std::endl;
            promptUserForStoragePath(fileManager, configManager);
        }
    }
    else {
        // Запрос пути у пользователя, если сохраненный путь не найден
        promptUserForStoragePath(fileManager, configManager);
    }
    // Загрузка последнего использованного ID
    Correspondence::nextId = fileManager.loadLastId();

    // Создание журналов
    Journal incomingJournal("incoming");
    Journal outgoingJournal("outgoing");

    // Загрузка данных из файлов
    fileManager.loadJournal(incomingJournal);
    fileManager.loadJournal(outgoingJournal);

    int choice;
    bool running = true;

    while (running) {
        displayMenu();
        std::cin >> choice;
        std::cin.ignore(); // Очистка буфера после ввода числа

        switch (choice) {
        case 1: { // Регистрация входящего документа
            std::cout << "\n=== Регистрация входящего документа ===\n";
            std::cout << "Для отмены и возврата в главное меню введите '0' в любом поле\n";

            std::string sender, topic;
            std::cout << "Отправитель: ";
            std::getline(std::cin, sender);

            if (sender == "0") {
                std::cout << "Регистрация отменена. Возврат в главное меню...\n";
                break;
            }

            std::cout << "Тема: ";
            std::getline(std::cin, topic);

            if (topic == "0") {
                std::cout << "Регистрация отменена. Возврат в главное меню...\n";
                break;
            }

            // Создаем документ
            Correspondence doc(getCurrentDate(), topic, sender, "incoming");

            // Спрашиваем о прикреплении файла
            std::cout << "Прикрепить PDF-файл? (1 - выбрать существующий, 0 - нет): ";
            int attachChoice;
            std::cin >> attachChoice;
            std::cin.ignore();

            std::string attachmentPath = "";

            if (attachChoice == 1) {
                std::cout << "Выберите PDF-файл для прикрепления...\n";
                attachmentPath = PdfManager::getFilePathFromDialog();

                if (!attachmentPath.empty()) {
                    try {
                        attachmentPath = fileManager.copyFileToStorage(attachmentPath, "incoming");
                        doc.setAttachmentPath(attachmentPath);
                        std::cout << "Файл успешно прикреплен.\n";
                    }
                    catch (const std::exception& e) {
                        std::cerr << "Ошибка при прикреплении файла: " << e.what() << std::endl;
                    }
                }
            }
            //else if (attachChoice == 2) { НЕРЕАЛИЗОВАНАЯ ФУНКЦИЯ СКАНИРОВАНИЯ ДОКУМЕНТА. ЗАДАЧА ВЫЯСНИТЬ РЕАЛИЗАЦИЮ ПОДКЛЮЧЕНИЯ СКАНЕРА ИЛИ ПОДТЯГИВАНИЯ ИЗ СПИСКА ЗАРЕГИСТРИРОВАННЫХ В ОС
            //    // Сканирование документа
            //    std::cout << "Подготовьте документ для сканирования...\n";
            //    attachmentPath = PdfManager::scanToPdf(fileManager.getStoragePath() + "/attachments");

            //    if (!attachmentPath.empty()) {
            //        doc.setAttachmentPath(attachmentPath);
            //        std::cout << "Документ успешно отсканирован и прикреплен.\n";
            //    }
            //    else {
            //        std::cout << "Ошибка при сканировании документа.\n";
            //    }
            //}

            incomingJournal.addCorrespondence(doc);
            fileManager.saveJournal(incomingJournal);
            std::cout << "Документ зарегистрирован с номером: " << doc.getId() << std::endl;
            break;
        }

        case 2: { // Регистрация исходящего документа
            std::cout << "\n=== Регистрация исходящего документа ===\n";
            std::cout << "Для отмены и возврата в главное меню введите '0' в любом поле\n";

            std::string recipient, topic;
            std::cout << "Получатель: ";
            std::getline(std::cin, recipient);

            if (recipient == "0") {
                std::cout << "Регистрация отменена. Возврат в главное меню...\n";
                break;
            }

            std::cout << "Тема: ";
            std::getline(std::cin, topic);

            if (topic == "0") {
                std::cout << "Регистрация отменена. Возврат в главное меню...\n";
                break;
            }

            // Создаем документ
            Correspondence doc(getCurrentDate(), topic, recipient, "outgoing");

            // Спрашиваем о прикреплении файла
            std::cout << "Прикрепить PDF-файл? (1 - выбрать существующий, 0 - нет): ";
            int attachChoice;
            std::cin >> attachChoice;
            std::cin.ignore();

            std::string attachmentPath = "";

            if (attachChoice == 1) {
                // Выбор существующего файла через диалог
                std::cout << "Выберите PDF-файл для прикрепления...\n";
                attachmentPath = PdfManager::getFilePathFromDialog();

                if (!attachmentPath.empty()) {
                    try {
                        attachmentPath = fileManager.copyFileToStorage(attachmentPath, "outgoing");
                        doc.setAttachmentPath(attachmentPath);
                        std::cout << "Файл успешно прикреплен.\n";
                    }
                    catch (const std::exception& e) {
                        std::cerr << "Ошибка при прикреплении файла: " << e.what() << std::endl;
                    }
                }
            }
            //else if (attachChoice == 2) {
            //    // Сканирование документа
            //    std::cout << "Подготовьте документ для сканирования...\n";
            //    std::string scanDir = fileManager.getStoragePath() + "/outgoing_attachments";
            //    attachmentPath = PdfManager::scanToPdf(scanDir);

            //    if (!attachmentPath.empty()) {
            //        doc.setAttachmentPath(attachmentPath);
            //        std::cout << "Документ успешно отсканирован и прикреплен.\n";
            //    }
            //    else {
            //        std::cout << "Ошибка при сканировании документа.\n";
            //    }
            //}

            outgoingJournal.addCorrespondence(doc);
            fileManager.saveJournal(outgoingJournal);
            std::cout << "Документ зарегистрирован с номером: " << doc.getId() << std::endl;
            break;
        }
        case 3: { // Просмотр журнала корреспонденции
            int journalChoice;
            std::cout << "\n=== Просмотр журнала ===\n";
            std::cout << "1. Входящая корреспонденция\n";
            std::cout << "2. Исходящая корреспонденция\n";
            std::cout << "0. Вернуться в главное меню\n"; // Добавлена опция возврата
            std::cout << "Выберите тип журнала: ";
            std::cin >> journalChoice;
            std::cin.ignore(); // Очистка буфера

            if (journalChoice == 0) {
                std::cout << "Возврат в главное меню...\n";
                break; // Выход из case 3
            }

            switch (journalChoice) {
            case 1:
                std::cout << "\n=== Журнал входящей корреспонденции ===\n";
                incomingJournal.displayJournal();
                break;
            case 2:
                std::cout << "\n=== Журнал исходящей корреспонденции ===\n";
                outgoingJournal.displayJournal();
                break;
            default:
                std::cout << "Неверный выбор журнала. Возврат в главное меню.\n";
            }
            break;
        }
        case 4: { // Поиск документа (изменен номер)
            std::string searchTerm;
            std::cout << "Введите строку для поиска: ";
            std::getline(std::cin, searchTerm);

            std::cout << "\n=== Результаты поиска ===\n";
            std::cout << "Входящие документы:\n";
            incomingJournal.searchCorrespondence(searchTerm);
            std::cout << "Исходящие документы:\n";
            outgoingJournal.searchCorrespondence(searchTerm);
            break;
        }
        case 5: { // Изменение места хранения (изменен номер)
            promptUserForStoragePath(fileManager, configManager);
            // Сохраняем журналы в новое место
            fileManager.saveJournal(incomingJournal);
            fileManager.saveJournal(outgoingJournal);
            break;
        }
        case 6: { // Экспорт журнала в TXT
            int journalType;
            std::cout << "\n=== Экспорт журнала ===\n";
            std::cout << "1. Экспортировать журнал входящей корреспонденции\n";
            std::cout << "2. Экспортировать журнал исходящей корреспонденции\n";
            std::cout << "0. Вернуться в главное меню\n";
            std::cout << "Выберите действие: ";
            std::cin >> journalType;
            std::cin.ignore();

            if (journalType == 0) {
                std::cout << "Возврат в главное меню...\n";
                break;
            }

            if (journalType != 1 && journalType != 2) {
                std::cout << "Неверный выбор журнала. Возврат в главное меню.\n";
                break;
            }

            std::string filename;
            std::cout << "Введите имя файла для экспорта (или 0 для отмены): ";
            std::getline(std::cin, filename);

            if (filename == "0") {
                std::cout << "Экспорт отменен. Возврат в главное меню...\n";
                break;
            }

            try {
                if (journalType == 1) {
                    fileManager.exportJournalToTXT(incomingJournal, filename);
                }
                else {
                    fileManager.exportJournalToTXT(outgoingJournal, filename);
                }
            }
            catch (const std::exception& e) {
                std::cerr << "Ошибка при экспорте: " << e.what() << std::endl;
            }
            break;
        }
        case 7: { // Открытие прикрепленного файла
            std::cout << "\n=== Открытие прикрепленного файла ===\n";
            std::cout << "1. Для входящего документа\n";
            std::cout << "2. Для исходящего документа\n";
            std::cout << "0. Вернуться в главное меню\n";
            std::cout << "Выберите тип документа: ";

            int docType;
            std::cin >> docType;
            std::cin.ignore();

            if (docType == 0) {
                std::cout << "Возврат в главное меню...\n";
                break;
            }

            Journal& journal = (docType == 1) ? incomingJournal : outgoingJournal;

            std::cout << "Введите номер документа: ";
            int docId;
            std::cin >> docId;
            std::cin.ignore();

            bool found = false;
            for (const auto& record : journal.getRecords()) {
                if (record.getId() == docId) {
                    found = true;
                    if (record.getAttachmentPath().empty()) {
                        std::cout << "К этому документу не прикреплен файл.\n";
                    }
                    else {
                        std::cout << "Открываю прикрепленный файл...\n";
                        if (!PdfManager::openPdfFile(record.getAttachmentPath())) {
                            std::cout << "Ошибка при открытии файла.\n";
                        }
                    }
                    break;
                }
            }

            if (!found) {
                std::cout << "Документ с указанным номером не найден.\n";
            }
            break;
        }
        case 0: // Выход
            running = false;
            std::cout << "Программа завершена.\n";
            break;
        default:
            std::cout << "Неверный выбор. Попробуйте снова.\n";
        }
    }
    // Перед выходом сохраняем последний ID
    fileManager.saveLastId(Correspondence::nextId);

    return 0;
}