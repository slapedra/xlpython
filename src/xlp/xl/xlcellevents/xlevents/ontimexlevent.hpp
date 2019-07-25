/*
 2008 Sebastien Lapedra
*/
#ifndef ontimexlevent_hpp
#define ontimexlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	#define SECS_PER_DAY (60.0 * 60.0 * 24.0)

	class OnTimeXLEventBuilder : public XLEventBuilder {
	public:
		class OnTimeXLEvent : public XLEvent {
		public:
			OnTimeXLEvent(const std::string& formula, const boost::shared_ptr<XLEvent>& handler)
			: xlpFormula_(formula, "", false, false), handler_(handler) {
				xlpFormula_.setToFree(true);
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				Xloper xlValue;
				Xloper xlFormula;
				Excel4(xlfFormulaConvert, &xlFormula, 3, xlpFormula_.get(), TempBool(false), TempBool(true));
				Excel4(xlfEvaluate, &xlValue, 1, &xlFormula);
				XlpOper xlpValue(&xlValue, "value");
				try {
					Real step  = xlpValue;
					if (step < 1.0) {
						step = XL_TIMESTEP;
					}
					Xloper& t = time(); 
					if (t->val.num >= lasttime_ + step / SECS_PER_DAY) {
						handler_->execute(xlCell);
						lasttime_ = t->val.num;
					}
				} catch (...) {
				}
				return true;
			}

		protected:
			XlpOper xlpFormula_;
			mutable Real lasttime_;
			boost::shared_ptr<XLEvent> handler_;
		};

		static std::string id() { return "ontime"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			++args.index;
			REQUIRE(args.index < args.args.size(), "not conform name");
			std::map< std::string, boost::shared_ptr<XLEventBuilder> >::const_iterator builderiter = 
				args.builders->find(args.args[args.index]);
			REQUIRE(builderiter != args.builders->end(), "unknow command: " << args.args[args.index]);
			std::list< boost::shared_ptr<XLEvent> > timeevents;
			builderiter->second->populate(args, xlDstRef, key, nameKey, src, timeevents, session);
			std::string namesrc = nameKey+"_args.time_"+key;
			boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+namesrc);
			xlInt.val.w = 32;
			Xloper xlAddr;
			Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlSrcRef->oper())));
			XlpOper xlpAddr(&xlAddr, "addr");
			std::string addr = xlpAddr;
			Size row = xlSrcRef->row()+1;
			Size col = xlSrcRef->col()+1;
			std::list< boost::shared_ptr<XLEvent> >::iterator iter = timeevents.begin();
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					std::string formula = "='"+addr+"'!R"+boost::lexical_cast<std::string>(row+k)
						+"C"+boost::lexical_cast<std::string>(col+j);	
					events.push_back(boost::shared_ptr<XLEvent>(new OnTimeXLEvent(formula, *iter)));
					++iter;
				}
			}
		}
	};
	
}

#endif

