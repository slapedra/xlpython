/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>

#include <xlp/xl/xlpoper.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcellevents/serievalue.hpp>

#include <iomanip>
#include <iostream>
#include <fstream>
#include <boost/archive/tmpdir.hpp>
#include <boost/archive/text_iarchive.hpp>
#include <boost/archive/text_oarchive.hpp>
#include <boost/serialization/base_object.hpp>
#include <boost/serialization/utility.hpp>
#include <boost/serialization/list.hpp>
#include <boost/serialization/assume_abstract.hpp>
#include <boost/serialization/vector.hpp>
#include <boost/serialization/variant.hpp>

namespace xlp {

	SerieValue::SerieCellMap SerieValue::serieCellMap_;

	SerieValue::SerieCellMap::iterator SerieValue::iterator() const {
		SerieCellMap::iterator iter = serieCellMap_.find(boost::make_tuple(
			xlCell_->row(), xlCell_->col(), xlCell_->idSheet()));
		return iter;
	}

	std::string SerieValue::workbook() const {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 32;
		Xloper xlAddr;
		Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlCell_->oper())));
		XlpOper xlpName(&xlAddr, "addr");
		std::string addr = xlpName;
		return addr;
	}

	void SerieValue::save(const Scalar& value) const {
		SerieCellMap::iterator iter = iterator();
		if (iter == serieCellMap_.end()) {
			SerieCellElem serie(xlCell_->row(), xlCell_->col(), xlCell_->idSheet(), workbook());
			serie.serie = boost::circular_buffer<Scalar>(size_, NAType());
			serie.serie.push_front(value);
			serieCellMap_.insert(serie);
		} else {
			if (iter->serie.capacity() != size_) {
				iter->serie.set_capacity(size_);
			}
			iter->serie.push_front(value);
		}
	}

	void SerieValue::del() const {
		SerieCellMap::iterator iter = iterator();
		if (iter == serieCellMap_.end()) {
			return;
		} 
		if (iter->serie.capacity() != size_) {
			iter->serie.set_capacity(size_);
		}
		if (iter->serie.size() > 0) {
			iter->serie.erase(iter->serie.begin());
		}
	}

	Scalar SerieValue::load(Size i) const {
		SerieCellMap::iterator iter = iterator();
		if (iter == serieCellMap_.end()) {
			return NullType();
		} 
		if (i<iter->serie.size()) {
			return iter->serie[i];
		} else {
			return NullType();
		}
	}
	
	void SerieValue::erase(const std::string& workbook) {
		SerieCellMap::iterator iter = serieCellMap_.begin();
		while (iter != serieCellMap_.end()) {
			if (iter->workbook == workbook) {
				iter = serieCellMap_.erase(iter);
			} else {
				++iter;
			}
		}
	}

	boost::shared_ptr< Range<Scalar> > SerieValue::range() const {
		SerieCellMap::iterator iter = iterator();
		if (iter == serieCellMap_.end()) {
			return boost::shared_ptr< Range<Scalar> >(new Range<Scalar>(0, 0));
		}  
		boost::shared_ptr< Range<Scalar> > data(new Range<Scalar>(1, iter->serie.size()));
		std::copy(iter->serie.begin(), iter->serie.end(), data->data.begin());
		return data;
	}

}

