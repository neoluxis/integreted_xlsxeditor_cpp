#pragma once

#include <QHash>
#include <QImage>
#include <QProgressBar>
#include <QSet>
#include <QString>
#include <QVector>
#include <QWidget>
#include <cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp>
#include <cc/neolux/utils/MiniXLSX/XLPictureReader.hpp>

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

class XLSXEditor : public QWidget {
    Q_OBJECT

public:
    explicit XLSXEditor(QWidget* parent = nullptr);
    ~XLSXEditor();

    void loadXLSX(const QString& filePath, const QString& sheetName, const QString& range);

private slots:
    void on_btnSave_clicked();
    void on_btnPreview_clicked();

protected:
    bool eventFilter(QObject* watched, QEvent* event) override;

private:
    Ui::XLSXEditor* ui;
    QString m_filePath;
    QString m_sheetName;
    QString m_range;
    QVector<DataEntry> m_data;  // 图片与描述及其位置
    QVector<DataItem*> m_dataItems;
    QHash<QString, int> m_indexByCell;
    QHash<QString, DataItem*> m_itemByCell;
    QSet<QString> m_dirtyCells;
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper* m_wrapper;
    cc::neolux::utils::MiniXLSX::XLPictureReader m_pictureReader;
    int m_sheetIndex;

    void parseRange(const QString& range, int& startRow, int& startCol, int& endRow, int& endCol);
    void loadData();
    void loadData(QProgressBar& progressBar);
    void displayData(bool previewOnly = false);
    bool saveData();
    int colToNum(const QString& col);
    QString numToCol(int num);
    QString cellKey(int row, int col) const;
    void showImageDialog(int row, int col);
    void clearDataItems();
    void resetState();
    void updateScrollWidgetSize();
    void syncPreviewButtonText();
    void syncPreviewVisibility();
    bool m_previewOnly;
    double m_itemScale;
};

}  // namespace xlsxeditor
}  // namespace fem
}  // namespace neolux
}  // namespace cc
