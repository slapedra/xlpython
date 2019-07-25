/*
 2008 Sebastien Lapedra
*/
#ifndef seekxlevent_hpp
#define seekxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class SeekXLEventBuilder : public XLEventBuilder {
	public:
		class SeekXLEvent : public XLEvent {
		public:
			SeekXLEvent(const boost::shared_ptr<XLRef>& xlTarget, bool goal) 
			: xlTarget_(xlTarget), goal_(goal) {
			}

			SeekXLEvent(const boost::shared_ptr<XLRef>& xlTarget, const boost::shared_ptr<XLRef>& xlDst) 
			: xlTarget_(xlTarget), xlDst_(xlDst), goal_(true) {
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				if (xlDst_.get() == NULL) {
					if (goal_) {
						Scalar value = valueCell(xlCell);
						XlpOper xlpValue(value, "value");
						xlpValue.setToFree(true);
						Excel4(xlcGoalSeek, 0, 3, &(*(xlTarget_->oper())), xlpValue.get(), &(*(xlCell->oper())));
					} else {
						static XLOPER xlValue;
						xlValue.xltype = xltypeNum;
						xlValue.val.num = 0.0;
						Excel4(xlcGoalSeek, 0, 3, &(*(xlTarget_->oper())), &xlValue, &(*(xlCell->oper())));
					}
				} else {
					Scalar value = valueCell(xlCell);
					XlpOper xlpValue(value, "value");
					xlpValue.setToFree(true);
					static XLOPER xlInt;
					xlInt.xltype = xltypeInt;
					xlInt.val.w = 32;
					Xloper xlAddr;
					Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlTarget_->oper())));
					XlpOper xlpAddr(&xlAddr, "addr");
					std::string addr = xlpAddr;
					Size row = xlTarget_->row()+1;
					Size col = xlTarget_->col()+1;
					std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row)
						+"C"+boost::lexical_cast<std::string>(col);	
					Excel4(xlcFormula, 0, 2, TempStrStl(formula), &(*(xlCell->oper())));
					if ( (!value.missing()) && (!value.error()) ) {
 						boost::shared_ptr<XLRef> xlDst2 = dstRef(xlDst_);
						Scalar valueDst = valueCell(xlDst2);
						XlpOper xlpValueDst(valueDst, "value");
						xlpValueDst.setToFree(true);
						Excel4(xlcFormula, 0, 2, xlpValueDst.get(), &(*(xlDst2->oper())));
						Excel4(xlcGoalSeek, 0, 3, &(*(xlTarget_->oper())), xlpValue.get(), &(*(xlDst2->oper())));
					}
				}
				return true;
			}

		protected:
			boost::shared_ptr<XLRef> xlTarget_;
			boost::shared_ptr<XLRef> xlDst_;
			bool goal_;
		};
		
		static std::string id() { return "seek"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+nameKey+"_args.src_"+key);
			Size row = xlSrcRef->row();
			Size col = xlSrcRef->col();
			try {
				boost::shared_ptr<XLRef> xlDstRef = session->checkOutput(src+"!"+nameKey+"_args.dst_"+key);
				Size drow = xlDstRef->row();
				Size dcol = xlDstRef->col();
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						boost::shared_ptr<XLRef> xlTarget(new XLRef(row+k, col+j, 1, 1, xlSrcRef->idSheet()));
						boost::shared_ptr<XLRef> xlDst(new XLRef(drow+k, dcol+j, 1, 1, xlDstRef->idSheet()));
						events.push_back(boost::shared_ptr<XLEvent>(new SeekXLEvent(xlTarget, xlDst)));
					}
				}
			} catch (...) {
				bool goal = false;
				if (args.index < args.args.size()) {
					if (args.args[args.index] == "goal") {
						goal = true;
						++args.index;
					}
				}
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						boost::shared_ptr<XLRef> xlTarget(new XLRef(row+k, col+j, 1, 1, xlSrcRef->idSheet()));
						events.push_back(boost::shared_ptr<XLEvent>(new SeekXLEvent(xlTarget, goal)));
					}
				}
			}
		}
	};
	
}

#endif

