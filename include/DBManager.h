#pragma once

#include <QString>
#include <QStringList>


/**=================================================================================================
 * Class: CDBManager
 * =================================================================================================
 * \brief DB���ݿ������.
 * \author @chensf
 * \date 2011-11-10 04:51:27
 *===============================================================================================**/
class Db;
class __declspec(dllexport) CDBManager
{
public:
	CDBManager();
	CDBManager(const QString& strDBName);

public:
	~CDBManager(void);

public:
	
	/**=================================================================================================
	 * Method: OpenDB
	 * =================================================================================================
	 * \brief �����ݿ�.
	 * \author @chensf
	 * \date 2011-11-10 04:42:33
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool OpenDB();


	/**=================================================================================================
	 * Method: OpenDB
	 * =================================================================================================
	 * \brief �����ݿ�.
	 * \author @chensf
	 * \date 2011-11-10 04:55:28
	 * \param strDBName Name of the string database.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool OpenDB(const QString& strDBName);
	
	
	/**=================================================================================================
	 * Method: CloseDB
	 * =================================================================================================
	 * \brief �ر����ݿ�.
	 * \author @chensf
	 * \date 2011-11-10 04:43:46
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool CloseDB();

	
	/**=================================================================================================
	 * Method: GetDataByKey
	 * =================================================================================================
	 * \brief ��ȡ��ֵ.
	 * \author @chensf
	 * \date 2011-11-10 04:44:39
	 * \param strKey The string key.
	 * \param [in,out] strData Information describing the string.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool GetDataByKey(const QString& strKey, QString& strData);


	/**=================================================================================================
	 * Method: GetDataByKey
	 * =================================================================================================
	 * \brief ��ȡ��ֵ.
	 * \author @chensf
	 * \date 2011-11-10 04:45:31
	 * \param strKey The string key.
	 * \param [in,out] pData If non-null, the data.
	 * \param [in,out] iSize Zero-based index of the size.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool    GetDataByKey(const QString& strKey, char* pData, int& iSize);

	
	/**=================================================================================================
	 * Method: SetDataByKey
	 * =================================================================================================
	 * \brief ���ü�ֵ��������ģ��������.
	 * \author @chensf
	 * \date 2011-11-10 04:45:50
	 * \param strKey The string key.
	 * \param strData Information describing the string.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool	SetDataByKey(const QString& strKey, const QString& strData);


	/**=================================================================================================
	 * Method: SetDataByKey
	 * =================================================================================================
	 * \brief ���ü�ֵ��������ģ��������.
	 * \author @chensf
	 * \date 2011-11-10 04:47:31
	 * \param strKey The string key.
	 * \param [in,out] pData If non-null, the data.
	 * \param iSize Zero-based index of the size.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool    SetDataByKey(const QString& strKey, char* pData, int iSize);

	
	/**=================================================================================================
	 * Method: DelDataByKey
	 * =================================================================================================
	 * \brief ɾ���ü�ֵ������.
	 * \author @chensf
	 * \date 2011-11-10 04:47:41
	 * \param strKey The string key.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool	DelDataByKey(const QString& strKey);


	/**=================================================================================================
	 * Method: GetAllKey
	 * =================================================================================================
	 * \brief ��ȡ���ݿ����м�ֵ.
	 * \author @chensf
	 * \date 2011-11-10 04:48:53
	 * \param [in,out] arrkeys The arrkeys.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool	GetAllKey(QStringList& arrkeys);

	
	/**=================================================================================================
	 * Method: GetAllKeyAndData
	 * =================================================================================================
	 * \brief ��ȡ���ݿ����м�ֵ������.
	 * \author @chensf
	 * \date 2011-11-10 04:49:14
	 * \param [in,out] mapData Information describing the map.
	 * \return true if it succeeds, false if it fails.
	 *===============================================================================================**/
	bool	GetAllKeyAndData(QMap<QByteArray,QByteArray>& mapData);

	
	/**=================================================================================================
	 * Method: IsHaveKey
	 * =================================================================================================
	 * \brief �ж��Ƿ���Ѿ��ڸü�ֵ.
	 * \author @chensf
	 * \date 2011-11-10 04:50:08
	 * \param strKey The string key.
	 * \return true if have key, false if not.
	 *===============================================================================================**/
	bool	IsHaveKey(const QString& strKey);
	

private:
	Db*		m_pDb;			
	QString m_strDBName;
};
