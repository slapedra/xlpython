/*
 2008 Sebastien Lapedra
*/
#ifndef serievalue_hpp
#define serievalue_hpp

#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlref.hpp>

#include <boost/multi_index_container.hpp>
#include <boost/multi_index/ordered_index.hpp>
#include <boost/multi_index/identity.hpp>
#include <boost/multi_index/member.hpp>
#include <boost/multi_index/composite_key.hpp>
#include <boost/circular_buffer.hpp>


namespace xlp {

	/*! SerieValue
	 */
	class SerieValue {
	public:
		SerieValue(const boost::shared_ptr<XLRef>& xlCell, Size size = 10)
		: xlCell_(xlCell), size_(size) {
		}

		void save(const Scalar& value) const;
		void del() const;
		Scalar load(Size i) const;
		boost::shared_ptr< Range<Scalar> > range() const;

		static void erase(const std::string& workbook);

	protected:
		struct SerieCellElem {
			Size row;
			Size col;
			long idsheet;
			std::string workbook;

			mutable boost::circular_buffer<Scalar> serie;
			
			SerieCellElem(Size srow, Size scol, long sidsheet, const std::string& sworkbook)
			: row(srow), col(scol), idsheet(sidsheet), workbook(sworkbook) {
			}
		};

		struct seriecellelem_key:composite_key<
			SerieCellElem,
			BOOST_MULTI_INDEX_MEMBER(SerieCellElem, Size, row),
			BOOST_MULTI_INDEX_MEMBER(SerieCellElem, Size, col), 
			BOOST_MULTI_INDEX_MEMBER(SerieCellElem, long, idsheet)
		>{};

		typedef multi_index_container<
			SerieCellElem, indexed_by<
				ordered_unique<seriecellelem_key> > 
		> SerieCellMap;

	protected:
		boost::shared_ptr<XLRef> xlCell_;
		Size size_;
		
		static SerieCellMap serieCellMap_;

		std::string workbook() const;
		SerieCellMap::iterator iterator() const;
	};
	
}

#endif

