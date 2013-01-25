// VisitData.h: interface for the CVisitData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_)
#define AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define		DB_TYPE_SQLSERVER	1
#define		DB_TYPE_ACCESS		2

class CVisitData  
{
public:
	CVisitData();
	virtual ~CVisitData();

private:
	_RecordsetPtr m_pDBRecordSet;	// 指向“数据库模块”的当前记录集的指针

	CString m_strDBServer;		//数据服务器名称
	CString m_strDBUserName;	//数据库连接用户名称
	CString m_strDBPassword;	//数据库连接密码
	CString m_strDBName;		//数据库名称
	int m_nDBType;				//数据库类型

protected:
	_ConnectionPtr m_pDBConn;		// 指向数据库连接对象的指针
	_StreamPtr m_pStm;				// 指向数据流对象的指针
	int m_nRecordCount;				// 当前记录集的记录数

public:
	//初始化数据库模块
	BOOL Initialize(LPCTSTR lpszDBServer, LPCTSTR lpszDBUserName, LPCTSTR lpszDBPassword, LPCTSTR lpszDBName, int nDBType = DB_TYPE_SQLSERVER);
	//反初始化数据库模块
	BOOL Uninitialize();
	
	BOOL BeginTrans();						//开始事务机制
	BOOL CommitTrans();						//提交事务
	BOOL RollbackTrans();					//事务执行失败，回滚

	BOOL Execute(LPCTSTR lpszSql);			//执行SQL命令
	int  InsertForID(LPCTSTR lpszSql);		//插入数据，并取回自动生成的ID

	BOOL SearchData(LPCTSTR lpszSqlSel);	//完成对数据的查询操作
	long GetRecordCount();					//获取查询结果的记录数目
	BOOL RecordSetMoveNext();				//将当前记录指针向后移动一条记录
	BOOL RecordSetMovePrev();				//将当前记录指针向前移动一条记录
	BOOL RecordSetMove(long lStepSize);		//移动当前记录指针
	BOOL RecordSetMoveLast();				//将最后一条记录设定为记录集的当前记录
	BOOL RecordSetMoveFirst();				//将第一条记录设定为记录集的当前记录
	BOOL IsRecordSetEOF();					//获取当前记录是否是记录集的最后一条记录
	BOOL IsRecordSetBOF();					//获取当前记录是否是记录集的第一条记录
	BOOL IsRecordSetEmpty();				//获取记录集是否为空
	_variant_t GetFieldValueOfRecordSet(LPCTSTR lpszFieldName);	//获取当前记录集中某个字段的值
	long GetFieldValue_Long(LPCTSTR lpszFieldName);				//获取当前记录集中某个字段的值（long）
	short GetFieldValue_Short(LPCTSTR lpszFieldName);			//获取当前记录集中某个字段的值（short）
	BYTE GetFieldValue_Byte(LPCTSTR lpszFieldName);				//获取当前记录集中某个字段的值（BYTE）
	BOOL GetFieldValue_Bool(LPCTSTR lpszFieldName);				//获取当前记录集中某个字段的值（BOOL）
	float GetFieldValue_Float(LPCTSTR lpszFieldName);			//获取当前记录集中某个字段的值（float）
	double GetFieldValue_Double(LPCTSTR lpszFieldName);			//获取当前记录集中某个字段的值（double）
	CString GetFieldValue_String(LPCTSTR lpszFieldName);		//获取当前记录集中某个字段的值（string）
	COleDateTime GetFieldValue_DateTime(LPCTSTR lpszFieldName);	//获取当前记录集中某个字段的值（COleDateTime）
	BOOL SetFieldValueOfRecordSet(LPCTSTR lpszFieldName, _variant_t varValue);	//设置当前记录集中某个字段的值
	BOOL SetFieldValue_Long(LPCTSTR lpszFieldName, long lValue);				//设置当前记录集中某个字段的值（long）
	BOOL SetFieldValue_Short(LPCTSTR lpszFieldName, short nValue);				//设置当前记录集中某个字段的值（short）
	BOOL SetFieldValue_Byte(LPCTSTR lpszFieldName, BYTE byValue);				//设置当前记录集中某个字段的值（BYTE）
	BOOL SetFieldValue_Bool(LPCTSTR lpszFieldName, BOOL bValue);				//设置当前记录集中某个字段的值（BOOL）
	BOOL SetFieldValue_Float(LPCTSTR lpszFieldName, float fValue);				//设置当前记录集中某个字段的值（float）
	BOOL SetFieldValue_Double(LPCTSTR lpszFieldName, double dValue);			//设置当前记录集中某个字段的值（double）
	BOOL SetFieldValue_String(LPCTSTR lpszFieldName, LPCTSTR lpszValue);		//设置当前记录集中某个字段的值（string）
	BOOL SetFieldValue_DateTime(LPCTSTR lpszFieldName, COleDateTime otDateTime);//设置当前记录集中某个字段的值（COleDateTime）
	BOOL UpdateRecordSet();					//更新当前记录集

protected:
};

#endif // !defined(AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_)
