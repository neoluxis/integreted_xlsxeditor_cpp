#ifndef DATAITEM_HPP
#define DATAITEM_HPP

#include <QWidget>
#include <QImage>
#include <QString>

namespace Ui {
class DataItem;
}

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
    void setRowCol(int row, int col);
    int getRow() const;
    int getCol() const;

signals:
    void deleteToggled(bool deleted);
    void imageClicked(int row, int col);

private slots:
    void on_btnMarkDel_clicked();

private:
    Ui::DataItem *ui;
    bool m_deleted;
    int m_row, m_col;
};

#endif // DATAITEM_HPP