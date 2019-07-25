/*
 2008 Sebastien Lapedra
*/
#ifndef dftxlevent_hpp
#define dftxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class DftXLEventBuilder : public XLEventBuilder {
	public:
		class DftXLEvent : public XLEvent {
		public:
			DftXLEvent(const std::string& formula) 
			: xlpFormula_(formula, "", false, false) {
				xlpFormula_.setToFree(true);
			}

			DftXLEvent(const std::string& formula, const boost::shared_ptr<XLRef>& xlDst)  
			: xlpFormula_(formula, "", false, false), xlDst_(xlDst) {
				xlpFormula_.setToFree(true);
			} 

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				Excel4(xlcFormula, 0, 2, xlpFormula_.get(), &(*(xlCell->oper())));
				return true;
			}

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				Scalar value = valueCell(xlCell);
				if (value.missing() || value.error()) {
					Excel4(xlcFormula, 0, 2, xlpFormula_.get(), &(*(xlCell->oper())));
				} else {
					if (xlDst_.get() != NULL) {
						Scalar value = valueCell(xlCell);
						XlpOper xlpValue(value, "value");
						xlpValue.setToFree(true);
						Excel4(xlcFormula, 0, 2, xlpValue.get(), &(*(dstRef(xlDst_)->oper())));
						Excel4(xlcFormula, 0, 2, xlpFormula_.get(), &(*(xlCell->oper())));
					}
				}
				return true;
			}

		protected:
			XlpOper xlpFormula_;
			boost::shared_ptr<XLRef> xlDst_;
		};
		
		static std::string id() { return "dft"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;	
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+nameKey+"_args.src_"+key);
			xlInt.val.w = 32;
			Xloper xlAddr;
			Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlSrcRef->oper())));
			XlpOper xlpAddr(&xlAddr, "addr");
			std::string addr = xlpAddr;
			Size row = xlSrcRef->row()+1;
			Size col = xlSrcRef->col()+1;
			try {
				boost::shared_ptr<XLRef> xlDstRef = session->checkOutput(src+"!"+nameKey+"_args.dst_"+key);
				Size drow = xlDstRef->row();
				Size dcol = xlDstRef->col();
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row+k)
							+"C"+boost::lexical_cast<std::string>(col+j);	
						boost::shared_ptr<XLRef> xlDst(new XLRef(drow+k, dcol+j, 1, 1, xlDstRef->idSheet()));
						events.push_back(boost::shared_ptr<XLEvent>(new DftXLEvent(formula, xlDst)));
					}
				}
			}
			catch (...) {
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row+k)
							+"C"+boost::lexical_cast<std::string>(col+j);	
						events.push_back(boost::shared_ptr<XLEvent>(new DftXLEvent(formula)));
					}
				}
			}
		}

	};
	
}

#endif

