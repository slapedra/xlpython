/*
 2008 Sebastien Lapedra
*/
#ifndef flipxlevent_hpp
#define flipxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class FlipXLEventBuilder : public XLEventBuilder {
	public:
		class FlipXLEvent : public XLEvent {
		public:
			FlipXLEvent(const std::string& falseFormula, const std::string& trueFormula, const XLSession* session)
			: xlpFalseFormula_(falseFormula, "", false, false), xlpTrueFormula_(trueFormula, "", false, false), 
				session_(session), falseformula_(new Range<Scalar>(1,1)) {
				xlpTrueFormula_.setToFree(true);
				xlpFalseFormula_.setToFree(true);
				falseformula_->data[0] = falseFormula;
			}

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				if ( (value.missing()) || (value.error()) || (value.scalar() == falsevalue.scalar()) ) {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
				} else {
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					static XLOPER xlBool;
					xlBool.xltype = xltypeBool;
					xlBool.val.boolean = true;
					Excel4(xlcEcho, 0, 1, &xlBool);
					xlBool.val.boolean = false;
					Excel4(xlcEcho, 0, 1, &xlBool);
					session_->formula(xlCell, falseformula_, false, session_->time() + XL_ACTION_WAITTIME);
				}
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
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
					static XLOPER xlBool;
					xlBool.xltype = xltypeBool;
					xlBool.val.boolean = true;
					Excel4(xlcEcho, 0, 1, &xlBool);
					xlBool.val.boolean = false;
					Excel4(xlcEcho, 0, 1, &xlBool);
					session_->formula(xlCell, falseformula_, false, session_->time() + XL_ACTION_WAITTIME);
				} else {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
				}
				return true;
			}

		protected:
			XlpOper xlpTrueFormula_;
			XlpOper xlpFalseFormula_;
			boost::shared_ptr< Range<Scalar> > falseformula_;
			const XLSession* session_;
		};
		
		static std::string id() { return "flip"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
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
						events.push_back(boost::shared_ptr<XLEvent>(new FlipXLEvent(falseformula, trueformula, session)));
					}
				}
			} catch(...) {
				try {
					boost::shared_ptr<XLRef> xlTrueSrcRef = session->checkOutput(src+"!"+nameKey+"_args.flip.true");
					boost::shared_ptr<XLRef>xlFalseSrcRef = session->checkOutput(src+"!"+nameKey+"_args.flip.false");
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
							events.push_back(boost::shared_ptr<XLEvent>(new FlipXLEvent(falseformula, trueformula, session)));
						}
					}
				} catch(...) {
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new FlipXLEvent("=false", "=true", session)));
						}
					}
				}
			}
		}
	};
	
}

#endif

