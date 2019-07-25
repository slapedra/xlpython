/*
 2008 Sebastien Lapedra
*/
#ifndef userserievalue_hpp
#define userserievalue_hpp

#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {

	/*! SerieValue
	 */
	class UserSerieValue {
	public:
		UserSerieValue(const boost::shared_ptr<XLRef>& xlCell, Size size = 10)
		: xlCell_(xlCell), size_(size) {
		}

		void save(const Scalar& value) const;
		void del() const;
		Scalar load(Size i) const;
		boost::shared_ptr< Range<Scalar> > range() const;

	protected:
		boost::shared_ptr<XLRef> xlCell_;
		Size size_;

		std::string idCellValue() const;
	};
	
}

#endif

