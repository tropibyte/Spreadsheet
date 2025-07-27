#pragma once

class CGrid32Mgr;

class CFormulator
{
	friend class CGrid32Mgr;

	static void SkipSpaces(const std::wstring& s, size_t& pos);
	static bool ParseCellRef(const std::wstring& token, UINT& row, UINT& col);
	static bool IsValidCellReference(const std::wstring& s, size_t& pos, size_t& row, size_t& col);
	static double GetCellValue(CGrid32Mgr* mgr, UINT row, UINT col,
		std::set<std::pair<UINT, UINT>>& visited);

	static double ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
		std::set<std::pair<UINT, UINT>>& visited);

	static double SumRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double MinMaxRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited, bool bMax);
	static double CountRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	static double ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
};
