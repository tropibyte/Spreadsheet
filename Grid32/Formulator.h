#pragma once

class CGrid32Mgr;
struct Grid32FormulaSession;	// per-evaluation state, private to Formulator.cpp

class CFormulator
{
	friend class CGrid32Mgr;

public:
	// What a formula evaluates to. Most formulas produce a number; one whose
	// outermost operation is a text function (LEFT, CONCAT, UPPER, ...)
	// produces a string, and `value` then holds that string read as a number
	// ("12" -> 12, "abc" -> 0) for callers that need one.
	struct Result
	{
		double value = 0.0;
		bool isText = false;
		std::wstring text;
	};

	// Evaluate formula source (no leading '='). `self` is the cell that holds
	// the formula, if any: it is on the evaluation path from the start, so a
	// reference back to it is a cycle exactly as it is when the same cell is
	// evaluated by position through EvaluateCell.
	static Result Evaluate(CGrid32Mgr* mgr, const std::wstring& formula,
		const GRIDCELL* self = nullptr);

	// Evaluate the formula held by (row, col). While a RecalcPass is open the
	// result is memoized, so each formula cell is evaluated once per pass no
	// matter how many formulas reference it.
	static Result EvaluateCell(CGrid32Mgr* mgr, UINT row, UINT col);

	// One recalculation pass: every EvaluateCell call made while it is alive
	// shares one memo. Open it, evaluate the formula cells, let it close.
	class RecalcPass
	{
	public:
		explicit RecalcPass(CGrid32Mgr* mgr);
		~RecalcPass();
		RecalcPass(const RecalcPass&) = delete;
		RecalcPass& operator=(const RecalcPass&) = delete;
	private:
		Grid32FormulaSession* m_pSession;
		Grid32FormulaSession* m_pOuter;
	};

	// How a numeric result reads under FMT_GENERAL.
	static std::wstring FormatGeneral(double value);

	// Formula-cell evaluations performed on this thread so far. Tests read the
	// difference across one recalculation to check the memoization.
	static size_t CellEvaluationCount();

	// Needed by file-scope text/date function helpers; safe to expose.
	// Accepts a *whole* reference token only ("A1"); anything with trailing
	// characters ("A1:B10", "A1B") is rejected so range syntax can't be
	// mistaken for a single cell. Syntax only: whether the cell lies inside a
	// particular grid is checked where the reference is evaluated.
	static bool ParseCellRef(const std::wstring& token, UINT& row, UINT& col);
	// Parse "A1" or "A1:B10" into an inclusive, normalized rectangular range.
	static bool ParseRange(const std::wstring& arg, UINT& sr, UINT& sc,
		UINT& er, UINT& ec);
private:
	struct Engine;	// evaluation internals, defined in Formulator.cpp

	static void SkipSpaces(const std::wstring& s, size_t& pos);
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
	static double AverageRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	// Additive level: term (('+'|'-') term)*. ParseExpression sits above this
	// and adds the comparison operators, which bind least tightly.
	static double ParseAdditive(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	static double ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
public:
	// Used by both internal helpers and (transitively) by file-scope text
	// function helpers — promoted so the namespace can call back in. Called
	// from outside an evaluation (as the tests do) it is an entry point in its
	// own right, memoized like Evaluate.
	static double EvalArg(CGrid32Mgr* mgr, const std::wstring& s, std::set<std::pair<UINT, UINT>>& visited);
};
