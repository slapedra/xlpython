/*
 2008 Sebastien Lapedra
*/
#ifndef xlevent_hpp
#define xlevent_hpp

#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlref.hpp>
#include <boost/algorithm/string/erase.hpp>

namespace xlp {

	class XLEvent {
	public:
		virtual inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const = 0;
		virtual bool entry(const boost::shared_ptr<XLRef>& xlCell) const { return execute(xlCell); }

		static Xloper& time(bool flag = false) {
			static Xloper xlNow;
			if (flag) {
				Excel4(xlfNow, &xlNow, 0);
			}
			return xlNow;
		}

		static long idSheet(long idd = 0) {
			static long id;
			if (idd != 0) {
				id = idd;
			}
			return id;
		}

	protected:
		inline Scalar valueCell(const boost::shared_ptr<XLRef>& xlCell) const {
			try {
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				xlInt.val.w = 5;
				Xloper xlValue;
				Excel4(xlfGetCell, &xlValue, 2, &xlInt, &(*(xlCell->oper())));
				XlpOper xlpValue(&xlValue, "value");
				Scalar value = xlpValue;
				return value;
			} catch (...) {
				return NullType();
			}
		}

		inline boost::shared_ptr<XLRef> dstRef(const boost::shared_ptr<XLRef>& xlCell) const {
			try {
				static XLOPER xlBool;
				static XLOPER xlInt;
				Xloper xlValue;
				xlInt.xltype = xltypeInt;
				xlBool.xltype = xltypeBool;
				xlBool.val.boolean = true;
				xlInt.val.w = 6;
				int result = Excel4(xlfGetCell, &xlValue, 2, &xlInt, &(*(xlCell->oper())));
				if (result != xlretSuccess) {
						return xlCell;
				} 
				XlpOper xlpValue(&xlValue, "value");
				std::string value = xlpValue;
				if ( value.empty() || (value[0] != '=') ) {
					return xlCell;
				} else {
					boost::erase_head(value, 1);
					XlpOper xlpId(value, "id");
					xlpId.setToFree(true);
					boost::shared_ptr<Xloper> xlRef(new Xloper());		
					int result = Excel4(xlfTextref, &(*xlRef), 2, xlpId.get(), &xlBool);
					if (result != xlretSuccess) {
						return xlCell;
					}  else {
						return boost::shared_ptr<XLRef>(new XLRef(xlRef));
					}
				}
			} catch (...) {
				return xlCell;
			}
		}

	};

	class XLEventBuilder {
	public:
		struct Args {
			Args() {
			}

			Args(const std::vector<std::string>& aargs, 
				Size iindex, std::map<std::string, boost::shared_ptr<XLEventBuilder> >* bbuilders) 
			: args(aargs), index(iindex), builders(bbuilders) {
			}

			std::vector<std::string> args;
			Size index;
			std::map<std::string, boost::shared_ptr<XLEventBuilder> > *builders;
		};

	public:
		virtual void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const = 0;
	};
	
}

#endif

