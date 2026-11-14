#pragma once

class CGrid32Mgr;
struct Grid32FormulaSession;	// per-evaluation state, private to Formulator.cpp

class CFormulator
{
	friend class CGrid32Mgr;

public:
	// What a formula evaluates to. Most formulas produce a number; one whose
	// outermost operation is a text function (LEFT, CONCAT, UPPER, ...), a
	// string literal or a reference to text produces a string, and `value`
	// then holds that string read as a number ("12" -> 12, "abc" -> 0) for
	// callers that need one.
	//
	// A formula that cannot be evaluated produces an error: `isError` is set,
	// `text` is its code (see the kError* constants) and `value` is NaN.
	//
	// A numeric result that is a date (DATE, TODAY, NOW, a date cell, a date
	// plus or minus a number, MIN/MAX of dates) sets `isDate`; `value` is its
	// serial, and a General cell shows it as a date.
	//
	// A logical result (a comparison, TRUE/FALSE, AND, OR, NOT, ISERROR,
	// ISBLANK, a TRUE/FALSE cell) sets `isBoolean`; `value` is 1 or 0, and the
	// cell shows TRUE or FALSE.
	struct Result
	{
		double value = 0.0;
		bool isText = false;
		bool isError = false;
		bool isDate = false;
		bool isBoolean = false;
		std::wstring text;
	};

	// The error codes a formula can produce.
	static constexpr const wchar_t* kErrorSyntax = L"#ERROR!";  // malformed formula
	static constexpr const wchar_t* kErrorName = L"#NAME?";     // unknown function or name
	static constexpr const wchar_t* kErrorDiv0 = L"#DIV/0!";    // division by zero
	static constexpr const wchar_t* kErrorNum = L"#NUM!";       // no finite result
	static constexpr const wchar_t* kErrorValue = L"#VALUE!";   // a value of the wrong kind ("abc"+1, VALUE("abc"))
	static constexpr const wchar_t* kErrorNA = L"#N/A";         // no value to give (IFS with no condition true)

	// IFERROR and ISERROR see every error but kErrorSyntax: a malformed
	// formula is a mistake in the formula, not a value it produced, and
	// catching it would hide the typo.

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

	// Reference syntax, public for the grid and the tools (GoTo, the console).
	// Accepts a *whole* reference token only ("A1", or absolute: "$A$1",
	// "$A1", "A$1"); anything with trailing
	// characters ("A1:B10", "A1B") is rejected so range syntax can't be
	// mistaken for a single cell. Syntax only: whether the cell lies inside a
	// particular grid is checked where the reference is evaluated.
	static bool ParseCellRef(const std::wstring& token, UINT& row, UINT& col);
	// Parse "A1" or "A1:B10" into an inclusive, normalized rectangular range.
	static bool ParseRange(const std::wstring& arg, UINT& sr, UINT& sc,
		UINT& er, UINT& ec);

	// Evaluate expression source (no leading '=') as a number: a text result
	// reads as the number it is (0 if none), an error as NaN. The tests use it;
	// the grid uses Evaluate and EvaluateCell. `visited` is the chain of
	// formula cells being evaluated (empty from outside).
	static double EvalArg(CGrid32Mgr* mgr, const std::wstring& s, std::set<std::pair<UINT, UINT>>& visited);

private:
	// The parser and everything it uses, defined in Formulator.cpp. It passes
	// typed values (number, text, blank, error; a number may be a date)
	// rather than a double plus thread-local side channels.
	struct Engine;
};
