/*
 2008 Sebastien Lapedra
*/
#ifndef partxlevent_hpp
#define partxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>
#include <boost/algorithm/string/find.hpp>

namespace xlp {

	#define MINHEIGHT			0.69

	class PartXLEventBuilder : public XLEventBuilder {
	public:
		class RPartXLEvent : public XLEvent {
		public:
			RPartXLEvent(const std::string& falseFormula, const std::string& trueFormula, 
				const boost::shared_ptr<XLRef>& xlSrcCell, const XLSession *session)
			: xlToHideCell_(xlSrcCell), xlpFalseFormula_(falseFormula, "", false, false), 
				xlpTrueFormula_(trueFormula, "", false, false), session_(session) {
				xlpTrueFormula_.setToFree(true);
				xlpFalseFormula_.setToFree(true);
			}

			inline static std::string id() {  return "row"; }

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.missing() || value.error() || (value.scalar() == falsevalue.scalar()) )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 1;
					Excel4(xlcRowHeight, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 2;
					Excel4(xlcRowHeight, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				}
				session_->unprotect(false);
				return true;
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.scalar() == falsevalue.scalar() )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 2;
					Excel4(xlcRowHeight, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 1;
					Excel4(xlcRowHeight, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				}
				session_->unprotect(false);
				return true;
			}

		protected:
			XlpOper xlpFalseFormula_;
			XlpOper xlpTrueFormula_;
			boost::shared_ptr<XLRef> xlToHideCell_;
			const XLSession *session_;
		};

		class CPartXLEvent : public XLEvent {
		public:
			CPartXLEvent(const std::string& falseFormula, const std::string& trueFormula, 
				const boost::shared_ptr<XLRef>& xlSrcCell, const XLSession *session)
			: xlToHideCell_(xlSrcCell), xlpFalseFormula_(falseFormula, "", false, false), 
				xlpTrueFormula_(trueFormula, "", false, false), session_(session) {
				xlpTrueFormula_.setToFree(true);
				xlpFalseFormula_.setToFree(true);
			}

			inline static std::string id() {  return "col"; }

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.missing() || value.error() || (value.scalar() == falsevalue.scalar()) )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 1;
					Excel4(xlcColumnWidth, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 2;
					Excel4(xlcColumnWidth, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				}
				session_->unprotect(false);
				return true;
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.scalar() == falsevalue.scalar() )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 2;
					Excel4(xlcColumnWidth, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
					xlInt.val.w = 1;
					Excel4(xlcColumnWidth, 0, 4, NULL, &(*(xlToHideCell_->oper())), NULL, &xlInt);
				}
				session_->unprotect(false);
				return true;
			}

		protected:
			XlpOper xlpFalseFormula_;
			XlpOper xlpTrueFormula_;
			boost::shared_ptr<XLRef> xlToHideCell_;
			const XLSession *session_;
		};
		
		class FPartXLEvent : public XLEvent {
		public:
			FPartXLEvent(const std::string& falseFormula, const std::string& trueFormula, 
				const boost::shared_ptr<XLRef>& xlSrcCell, const XLSession *session)
			: xlToHideCell_(xlSrcCell), xlpFalseFormula_(falseFormula, "", false, false), 
				xlpTrueFormula_(trueFormula, "", false, false), session_(session) {
				xlpTrueFormula_.setToFree(true);
				xlpFalseFormula_.setToFree(true);
			}

			inline static std::string id() {  return "fast"; }

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.missing() || value.error() || (value.scalar() == falsevalue.scalar()) )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
				Excel4(xlcRowHeight, 0, 2, TempNum(MINHEIGHT), &(*(xlToHideCell_->oper())));
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					Excel4(xlcRowHeight, 0, 3, NULL, &(*(xlToHideCell_->oper())), TempBool(true));
				}
				session_->unprotect(false);
				return true;
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( value.scalar() == falsevalue.scalar() )  {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					Excel4(xlcRowHeight, 0, 3, NULL, &(*(xlToHideCell_->oper())), TempBool(true));
				} else {
					session_->unprotect(true);
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
					Excel4(xlcRowHeight, 0, 2, TempNum(MINHEIGHT), &(*(xlToHideCell_->oper())));
				}
				session_->unprotect(false);
				return true;
			}

		protected:
			XlpOper xlpFalseFormula_;
			XlpOper xlpTrueFormula_;
			boost::shared_ptr<XLRef> xlToHideCell_;
			const XLSession *session_;
		};

		static std::string id() { return "part"; }


		inline void populate(
			Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			using namespace boost;
			++args.index;
			REQUIRE(args.index < args.args.size(), "not conform name");
			if (args.args[args.index] == "row") {
				populatePart<RPartXLEvent>(xlDstRef, key, nameKey, src, events, session);
			} else if (args.args[args.index] == "col") {
				populatePart<CPartXLEvent>(xlDstRef, key, nameKey, src, events, session);
			} else if (args.args[args.index] == "fast") {
				populatePart<FPartXLEvent>(xlDstRef, key, nameKey, src, events, session);
			} else {
				FAIL ("not conform name");
			}
		}

		template <class PartClass>
		inline void populatePart(const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			boost::shared_ptr<XLRef> xlSrcCell = session->checkOutput(src+"!"+nameKey+"_args.src_"+key);
			try {
				boost::shared_ptr<XLRef> xlTrueSrcRef = session->checkOutput(src+"!"+nameKey+"_args.true_"+key);
				boost::shared_ptr<XLRef>xlFalseSrcRef = session->checkOutput(src+"!"+nameKey+"_args.false_"+key);
				xlInt.val.w = 32;
				Xloper xlTrueAddr;
				Xloper xlFalseAddr;
				Excel4(xlfGetCell, &xlTrueAddr, 2, &xlInt, &(*(xlTrueSrcRef->oper())));
				Excel4(xlfGetCell, &xlFalseAddr, 2, &xlInt, &(*(xlFalseSrcRef->oper())));
				XlpOper xlpTrueAddr(&xlTrueAddr, "addr");
				XlpOper xlpFalseAddr(&xlFalseAddr, "addr");
				std::string trueaddr = xlpTrueAddr;
				std::string falseaddr = xlpFalseAddr;
				Size truerow = xlTrueSrcRef->row()+1;
				Size truecol = xlTrueSrcRef->col()+1;
				Size falserow = xlFalseSrcRef->row()+1;
				Size falsecol = xlFalseSrcRef->col()+1;
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						std::string trueformula = "='"+trueaddr+"'!R"+boost::lexical_cast<std::string>(truerow+k)
							+"C"+boost::lexical_cast<std::string>(truecol+j);	
						std::string falseformula = "='"+falseaddr+"'!R"+boost::lexical_cast<std::string>(falserow+k)
							+"C"+boost::lexical_cast<std::string>(falsecol+j);		
						events.push_back(boost::shared_ptr<XLEvent>(new PartClass(
							falseformula, trueformula, xlSrcCell, session)));
					}
				}
			} catch(...) {
				try {
					boost::shared_ptr<XLRef> xlTrueSrcRef = session->checkOutput(
						src+"!"+nameKey+"_args.part."+PartClass::id()+".true");
					boost::shared_ptr<XLRef>xlFalseSrcRef = session->checkOutput(
						src+"!"+nameKey+"_args.part"+PartClass::id()+".false");
					xlInt.val.w = 32;
					Xloper xlTrueAddr;
					Xloper xlFalseAddr;
					Excel4(xlfGetCell, &xlTrueAddr, 2, &xlInt, &(*(xlTrueSrcRef->oper())));
					Excel4(xlfGetCell, &xlFalseAddr, 2, &xlInt, &(*(xlFalseSrcRef->oper())));
					XlpOper xlpTrueAddr(&xlTrueAddr, "addr");
					XlpOper xlpFalseAddr(&xlFalseAddr, "addr");
					std::string trueaddr = xlpTrueAddr;
					std::string falseaddr = xlpFalseAddr;
					Size truerow = xlTrueSrcRef->row()+1;
					Size truecol = xlTrueSrcRef->col()+1;
					Size falserow = xlFalseSrcRef->row()+1;
					Size falsecol = xlFalseSrcRef->col()+1;
					std::string trueformula = "='"+trueaddr+"'!R"+boost::lexical_cast<std::string>(truerow)
						+"C"+boost::lexical_cast<std::string>(truecol);	
					std::string falseformula = "='"+falseaddr+"'!R"+boost::lexical_cast<std::string>(falserow)
						+"C"+boost::lexical_cast<std::string>(falsecol);		
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new PartClass(
								falseformula, trueformula, xlSrcCell, session)));
						}
					}
				} catch(...) {
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new PartClass(
								"=false", "=true", xlSrcCell, session)));
						}
					}
				}
			}
		}
	};
	
}

#endif

