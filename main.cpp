#include <iostream>
#include <string>
#include <vector>
#include <sstream>
#include <filesystem>
#include "OpenXLSX.hpp"

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
    if (argc < 3 || argc > 4)
    {
        std::cerr << "Uso: " << argv[0] << " <arquivo.xlsx> <titulo da coluna> [linha do título]\n";
        return 1;
    }

    std::string filePath = argv[1];
    std::string columnTitle = argv[2];
    int titleRow = (argc == 4) ? std::stoi(argv[3]) : 1;

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
        auto sheet = doc.workbook().worksheet((argc > 3) ? argv[3] : sheetNames[0]);

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
            std::cerr << "Coluna '" << columnTitle << "' não encontrada.\n";
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

        doc.close();
    }
    catch (const std::exception &e)
    {
        std::cerr << "Erro: " << e.what() << "\n";
        return 1;
    }

    return 0;
}
