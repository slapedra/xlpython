/*
 2008 Sebastien Lapedra
*/
#ifndef onexlevent_hpp
#define onexlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class OneXLEventBuilder : public XLEventBuilder {
	public:
		class OneXLEvent : public XLEvent {
		public:
			OneXLEvent(const std::string& falseFormula, const std::string& trueFormula, 
				const std::vector< boost::shared_ptr<XLRef> >& xlRefs)
			: xlpFalseFormula_(falseFormula, "", false, false), 
				xlpTrueFormula_(trueFormula, "", false, false), xlRefs_(xlRefs) {
				xlpTrueFormula_.setToFree(true);
				xlpFalseFormula_.setToFree(true);
			}

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				Scalar value = valueCell(xlCell);
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFalseFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				Scalar falsevalue = xlpValue;
				for (Size i=0;i<xlRefs_.size();++i) {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlRefs_[i]->oper())));
				}
				if ( (value.missing()) || (value.error()) || (value.scalar() == falsevalue.scalar()) ) {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
				} else {
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
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
				for (Size i=0;i<xlRefs_.size();++i) {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlRefs_[i]->oper())));
				}
				if ( value.scalar() == falsevalue.scalar() )  {
					Excel4(xlcFormula, 0, 2, xlpTrueFormula_.get(), &(*(xlCell->oper())));
				} else {
					Excel4(xlcFormula, 0, 2, xlpFalseFormula_.get(), &(*(xlCell->oper())));
				}
				return true;
			}

		protected:
			XlpOper xlpTrueFormula_;
			XlpOper xlpFalseFormula_;
			std::vector< boost::shared_ptr<XLRef> > xlRefs_;
		};
		
		static std::string id() { return "one"; }

		inline void populate(Args& args,
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			std::vector< boost::shared_ptr<XLRef> > xlRefs;
			Size drow = xlDstRef->row();
			Size dcol = xlDstRef->col();
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					boost::shared_ptr<XLRef> xlRef(new XLRef(drow+k, dcol+j, 1, 1, xlDstRef->idSheet()));
					xlRefs.push_back(xlRef);
				}
			}
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
						events.push_back(boost::shared_ptr<XLEvent>(new OneXLEvent(falseformula, trueformula, xlRefs)));
					}
				}
			} catch(...) {
				try {
					boost::shared_ptr<XLRef> xlTrueSrcRef = session->checkOutput(src+"!"+nameKey+"_args.one.true");
					boost::shared_ptr<XLRef>xlFalseSrcRef = session->checkOutput(src+"!"+nameKey+"_args.one.false");
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
							events.push_back(boost::shared_ptr<XLEvent>(new OneXLEvent(falseformula, trueformula, xlRefs)));
						}
					}
				} catch(...) {
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new OneXLEvent("=false", "=true", xlRefs)));
						}
					}
				}
			}
		}
	};
	
}

#endif

