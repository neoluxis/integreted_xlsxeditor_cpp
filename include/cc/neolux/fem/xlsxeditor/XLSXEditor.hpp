#pragma once

#include <QWidget>
#include <QString>
#include <QVector>
#include <QProgressDialog>
#include <cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp>

namespace Ui {
class XLSXEditor;
}

namespace cc {
namespace neolux {
namespace fem {
namespace xlsxeditor {

struct DataEntry {
    int row, col;
    QImage image;
    QString desc;
    bool deleted;
};

class DataItem;

class XLSXEditor : public QWidget
{
    Q_OBJECT

public:
    explicit XLSXEditor(QWidget *parent = nullptr);
    ~XLSXEditor();

    void loadXLSX(const QString &filePath, const QString &sheetName, const QString &range);

private slots:
    void on_btnSave_clicked();
    void on_btnRestore_clicked();

private:

    Ui::XLSXEditor *ui;
    QString m_filePath;
    QString m_sheetName;
    QString m_range;
    QVector<DataEntry> m_data; // image and description with position
    QVector<DataItem*> m_dataItems;
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper *m_wrapper;
    int m_sheetIndex;

    void parseRange(const QString &range, int &startRow, int &startCol, int &endRow, int &endCol);
    void loadData();
    void loadData(QProgressDialog &progress);
    void displayData();
    void saveData();
    void restoreData();
    int colToNum(const QString &col);
    QString numToCol(int num);
    void showImageDialog(int row, int col);
};

} // namespace xlsxeditor
} // namespace fem
} // namespace neolux
} // namespace cc
