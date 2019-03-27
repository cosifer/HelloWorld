#ifndef EXCELPROCESS_H
#define EXCELPROCESS_H

#include <QtCore/qglobal.h>
#include "ISheetT.h"
#include "libxl.h"

#include <QMap>
#include <QDebug>
#include <QObject>
#include <conio.h>
#include <iostream>
#include <windows.h>
#include <QStringList>

#include "include.h"


using namespace libxl;

#ifdef EXCELPROCESS_LIB
# define EXCELPROCESS_EXPORT Q_DECL_EXPORT
#else
# define EXCELPROCESS_EXPORT Q_DECL_IMPORT
#endif

class EXCELPROCESS_EXPORT ExcelProcess : public QObject
{
public:
	ExcelProcess(QObject *parent);
	~ExcelProcess();

	////获取Excel文件数据
	//bool FileData(const QString strFileName, QStringList& arrField, QList<QStringList>& arrFileData, QString strFlag = QString::fromLocal8Bit(""));

	//获取Excel文件数据
	bool FileData(const QString strFileName, QStringList& arrField, QList<QStringList>& arrFileData, QString strFlag = QString::fromLocal8Bit(""));

	//保存数据到文件
	bool SaveToFile(const QMap<QString, QStringList>& arrData, QString strFileName);
	bool SaveToFile(QList<ProjInfo> arrProInfo, const QString strFileName);
	bool SaveToFile(ProjInfo info, QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ);

	//导出查询项目汇总信息
	bool SaveToFile(QStringList arrField, QList<QStringList> arrFileData, Book* book, Sheet* sheet, const QList<int> arrNumIndex, QList<double> arrZJ, int iSize);
	//bool SaveToFile(QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ, int iSize);
	bool SaveToFile(QList<QStringList> arr, const QString strFileName, QList<int> arrNumIndex, int iSheetIndex);

	//导出ERP本次处理结果数据
	bool SaveToFile(QList<QStringList> arrData, QString strFileName);

	//导出查询项目的铁附件汇总数据
	bool SaveToFile(QString strFileName, QList<QStringList> arrData1, QList<QStringList> arrData2);

	//将表格中的数据进行分析，将不同地方用颜色进行提示
	void ExcelDataCompare(const QString strFileName);

private:

private:
	QMap<QString, QStringList> m_arrFileData;
	QList<QString> m_WZFlags;
	QList<QString> m_WZBFlags;

	//检测是否有错误字符的列名称
	QList<QString> m_XMCheck;
	QList<QString> m_WZCheck;
	QList<QString> m_SGCheck;
	QList<QString> m_JSCheck;
	QList<QString> m_WZBCheck;
};

#endif // EXCELPROCESS_H
