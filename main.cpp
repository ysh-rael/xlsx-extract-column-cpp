#include <iostream>
#include <string>
#include <vector>
#include <sstream>
#include <filesystem>
#include "OpenXLSX.hpp"

#ifdef _WIN32
#include <windows.h>
#endif

using namespace OpenXLSX;
namespace fs = std::filesystem;

std::string joinCSV(const std::vector<std::string> &values)
{
    std::ostringstream oss;
    for (size_t i = 0; i < values.size(); ++i)
    {
        if (i > 0)
            oss << ",";
        oss << values[i];
    }
    return oss.str();
}

int main(int argc, char *argv[])
{
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
#endif

    if (argc < 3 || argc > 5)
    {
        std::cerr << "Uso: " << argv[0] << " <arquivo.xlsx> <titulo da coluna> [linha do título] [nome da planilha]\n";
        return 1;
    }

    std::string filePath = argv[1];
    std::string columnTitle = argv[2];
    int titleRow = 1;
    std::string sheetName;

    if (argc >= 4)
    {
        titleRow = std::stoi(argv[3]);
    }

    if (!fs::exists(filePath))
    {
        std::cerr << "Arquivo não encontrado: " << filePath << "\n";
        return 1;
    }

    try
    {
        XLDocument doc;
        doc.open(filePath);

        auto sheetNames = doc.workbook().worksheetNames();
        if (argc == 5)
        {
            sheetName = argv[4];
        }
        else
        {
            sheetName = sheetNames[0]; // pega a primeira planilha
        }

        auto sheet = doc.workbook().worksheet(sheetName);

        // Encontrar o índice da coluna com o título desejado
        int columnIndex = -1;
        int col = 1;
        while (true)
        {
            auto cell = sheet.cell(XLCellReference(titleRow, col));
            if (cell.value().type() == XLValueType::Empty)
                break;
            if (cell.value().get<std::string>() == columnTitle)
            {
                columnIndex = col;
                break;
            }
            ++col;
        }

        if (columnIndex == -1)
        {
            std::cerr << "Coluna '" << columnTitle << "' não encontrada na linha " << titleRow << ".\n";
            return 1;
        }

        // Coletar os valores abaixo do título
        std::vector<std::string> values;
        int row = titleRow + 1;
        while (true)
        {
            auto cell = sheet.cell(XLCellReference(row, columnIndex));
            if (cell.value().type() == XLValueType::Empty)
                break;
            values.push_back(cell.value().get<std::string>());
            ++row;
        }

        std::cout << joinCSV(values) << std::endl;

    }
    catch (const std::exception &e)
    {
        std::cerr << "Erro: " << e.what() << "\n";
        return 1;
    }

    return 0;
}
