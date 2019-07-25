/*
 2008 Sebastien Lapedra
*/
#ifndef seriexlevent_hpp
#define seriexlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>
#include <xlp/xl/xlcellevents/serievalue.hpp>
#include <xlp/xl/xlcellevents/userserievalue.hpp>
#include <boost/algorithm/string.hpp> 
#include <boost/algorithm/string/trim.hpp>

namespace xlp {

	class SerieXLEventBuilder : public XLEventBuilder {
	public:
		class SerieXLEvent : public XLEvent {
		public:
			SerieXLEvent(Size size)
			: size_(size), dft_(false), xlpFormula_("", "", false, false) {
				xlpFormula_.setToFree(true);
			}

			SerieXLEvent(const std::string& formula, Size size)
			: size_(size), dft_(true), xlpFormula_(formula, "", false, false) {
				xlpFormula_.setToFree(true);
			}

			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				SerieValue serie(xlCell, size_);
				Scalar value = valueCell(xlCell);
				if ( (!value.missing()) && (!value.error()) ) {
					serie.save(value);
					XlpOper xlpValue(value, "value");
					xlpValue.setToFree(true);
					Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(xlCell->oper())));
					Excel4(xlcFormula, 0, 2, &(*xlpValue.get()), &(*(xlCell->oper())));
				} else {
					serie.del();
					if (!dft_) { 
						XlpOper xlpValue(serie.load(0), "value");
						xlpValue.setToFree(true);
						Excel4(xlcFormula, 0, 2, &(*xlpValue.get()), &(*(xlCell->oper())));
					} else {
						Scalar value = serie.load(0);
						if (value.missing() || value.error()) {
							Excel4(xlcFormula, 0, 2, xlpFormula_.get(), &(*(xlCell->oper())));
						} else {
							XlpOper xlpValue(value, "value");
							xlpValue.setToFree(true);
							Excel4(xlcFormula, 0, 2, &(*xlpValue.get()), &(*(xlCell->oper())));
						}
					}
				}
				return true;
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				return false;
			}

		protected:
			Size size_;
			bool dft_;
			XlpOper xlpFormula_;
		};

		class UserSerieXLEvent : public XLEvent {
		public:
			UserSerieXLEvent(Size size)
			: size_(size) {
			}


			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				UserSerieValue serie(xlCell, size_);
				Scalar value = valueCell(xlCell);
				if ( (!value.missing()) && (!value.error()) ) {
					serie.save(value);
					XlpOper xlpValue(value, "value");
					xlpValue.setToFree(true);
					Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(xlCell->oper())));
					Excel4(xlcFormula, 0, 2, &(*xlpValue.get()), &(*(xlCell->oper())));
				} else {
					if (!(value.scalar() == Scalar("xlp_last").scalar())) {
						serie.del();
					} 
					XlpOper xlpValue(serie.load(0), "value");
					xlpValue.setToFree(true);
					Excel4(xlcFormula, 0, 2, &(*xlpValue.get()), &(*(xlCell->oper())));
				}
				return true;
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const {
				return false;
			}

		protected:
			Size size_;
		};
		
		static std::string id() { return "serie"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			REQUIRE(args.index < args.args.size(), "not conform name");
			bool flag = true;
			if (args.args[args.index] == "user") { 
				++args.index;
				REQUIRE(args.index < args.args.size(), "not conform name");
				flag = false;
			}
			Size size = boost::lexical_cast<Size>(args.args[args.index]);
			if (flag) {
				populateSerie(xlDstRef, key, nameKey, src, events, session, size); 
			} else {
				populateUserSerie(xlDstRef, key, nameKey, src, events, session, size);
			}
		}

		inline void populateSerie(const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session, Size size) const {
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			try {
				boost::shared_ptr<XLRef> xlSrcRef = session->checkOutput(src+"!"+nameKey+"_args.src_"+key);
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
						events.push_back(boost::shared_ptr<XLEvent>(new SerieXLEvent(formula, size)));
					}
				}
			}	catch (...) {
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						events.push_back(boost::shared_ptr<XLEvent>(new SerieXLEvent(size)));
					}
				}
			}
		}

		inline void populateUserSerie(const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session, Size size) const {
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new UserSerieXLEvent(size)));
				}
			}
		}

	};
	
}

#endif

