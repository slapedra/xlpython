/*
 2008 Sebastien Lapedra
*/


#ifndef xlcellevent_hpp
#define xlcellevent_hpp

#define NOMINMAX
#include <windows.h>
#include <xlp/xl/xldefines.hpp>
#include <xlp/xl/xlcall.h>
#include <xlp/xl/framewrk.hpp>
#include <xlp/xl/xlpoper.hpp>
#include <xlp/util/singleton.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>
#include <sstream>
#include <stdlib.h> 
#include <xlp/xl/xloper.hpp>
#include <list>
#include <set>
#include <boost/timer.hpp>
#include <boost/function.hpp>
#include <boost/multi_index_container.hpp>
#include <boost/multi_index/ordered_index.hpp>
#include <boost/multi_index/identity.hpp>
#include <boost/multi_index/member.hpp>
#include <boost/multi_index/composite_key.hpp>
#include <boost/tuple/tuple.hpp>

using boost::multi_index_container;
using namespace boost::multi_index;

namespace xlp {
 
	/*! XLCellEvent
	 */
	class XLCellEvent {
	public:
		inline virtual bool onEntry() const { return false; }
		inline virtual bool onDC() const { return false; }
		inline virtual bool onTime() const { return false; }
		inline virtual bool onOpen() const { return false; }
		inline virtual bool onSheet() const { return false; }
		inline virtual bool onSKey() const { return false; }

		static Xloper& time(bool flag = false) {
			static Xloper xlNow;
			if (flag) {
				Excel4(xlfNow, &xlNow, 0);
			}
			return xlNow;
		}
	};

	/*! XLCellEventMap
	 */	
	class XLCellEventMap {
	protected:
		struct CellEventElem {
			Size row;
			Size col;
			long idsheet;
			boost::shared_ptr<XLCellEvent> handler;
			
			CellEventElem(Size srow, Size scol, long sidsheet, 
				const boost::shared_ptr<XLCellEvent>& shandler)
			: row(srow), col(scol), idsheet(sidsheet), handler(shandler) {
			}
		};

		struct celleventelem_key:composite_key<
			CellEventElem,
			BOOST_MULTI_INDEX_MEMBER(CellEventElem, Size, row),
			BOOST_MULTI_INDEX_MEMBER(CellEventElem, Size, col), 
			BOOST_MULTI_INDEX_MEMBER(CellEventElem, long, idsheet)
		>{};

		typedef multi_index_container<
			CellEventElem, indexed_by<
				ordered_unique<celleventelem_key> > 
		> CellEventMap;

	public:
		typedef CellEventMap::const_iterator const_iterator;

	public:
		inline void add(Size row, Size col, long idsheet, 
			const boost::shared_ptr<XLCellEvent>& handler) const {
			celleventmap_.insert(CellEventElem(row, col, idsheet, handler));
		}

		inline boost::shared_ptr<XLCellEvent> handler(Size row, Size col, long idsheet) const {
			CellEventMap::const_iterator iter = celleventmap_.find(boost::make_tuple(row, col, idsheet));
			if (iter != celleventmap_.end()) {
				return iter->handler;
			} else {
				return boost::shared_ptr<XLCellEvent>();
			}
		}

		inline const_iterator begin() const { return celleventmap_.begin(); }
		inline const_iterator end() const { return celleventmap_.end(); }

	protected:
		mutable CellEventMap celleventmap_;
	};
	
}

#endif
