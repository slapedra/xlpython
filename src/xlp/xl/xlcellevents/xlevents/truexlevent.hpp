/*
 2008 Sebastien Lapedra
*/
#ifndef truexlevent_hpp
#define truexlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class TrueXLEventBuilder : public XLEventBuilder {
	public:
		class TrueXLEvent : public XLEvent {
		public:
			TrueXLEvent(const std::string& formula) 
			: xlpFormula_(formula, "", false, false) {
				xlpFormula_.setToFree(true);
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(xlCell->oper())));
				Excel4(xlcFormula, 0, 2, xlpFormula_.get(), &(*(xlCell->oper())));
				return true;
			}

		protected:
			XlpOper xlpFormula_;
		};
		
		static std::string id() { return "true"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			try {
				boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+nameKey+"_args.true_"+key);
				xlInt.val.w = 32;
				Xloper xlAddr;
				Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlSrcRef->oper())));
				XlpOper xlpAddr(&xlAddr, "addr");
				std::string addr = xlpAddr;
				Size row = xlSrcRef->row()+1;
				Size col = xlSrcRef->col()+1;
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row+k)
									+"C"+boost::lexical_cast<std::string>(col+j);	
						events.push_back(boost::shared_ptr<XLEvent>(new TrueXLEvent(formula)));
					}
				}
			} catch(...) {
				try {
					boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+nameKey+"_args.true");
					xlInt.val.w = 32;
					Xloper xlAddr;
					Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlSrcRef->oper())));
					XlpOper xlpAddr(&xlAddr, "addr");
					std::string addr = xlpAddr;
					Size row = xlSrcRef->row()+1;
					Size col = xlSrcRef->col()+1;
					std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row)
							+"C"+boost::lexical_cast<std::string>(col);	
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new TrueXLEvent(formula)));
						}
					}
				} catch(...) {
					for (Size k=0;k<xlDstRef->rows();++k) {
						for (Size j=0;j<xlDstRef->cols();++j) {
							events.push_back(boost::shared_ptr<XLEvent>(new TrueXLEvent("=true")));
						}
					}
				}
			}
		}
	};
	
}

#endif

