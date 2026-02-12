#pragma once

#include <QWidget>
#include <QImage>
#include <QString>

namespace Ui {
class DataItem;
}

namespace cc {
namespace neolux {
namespace fem {
namespace xlsxeditor {

class DataItem : public QWidget
{
    Q_OBJECT

public:
    explicit DataItem(QWidget *parent = nullptr);
    ~DataItem();

    void setImage(const QImage &image);
    void setDescription(const QString &desc);
    QString getDescription() const;
    void setDeleted(bool deleted);
    bool isDeleted() const;
    void setPosition(int row, int col);
    void setRowCol(int row, int col);
    int getRow() const;
    int getCol() const;

signals:
    void imageClicked(int row, int col);
    void deleteToggled(bool deleted);

private slots:
    void on_btnMarkDel_clicked();

private:
    Ui::DataItem *ui;
    bool m_deleted;
    int m_row, m_col;
};

} // 命名空间 xlsxeditor
} // 命名空间 fem
} // 命名空间 neolux
} // 命名空间 cc