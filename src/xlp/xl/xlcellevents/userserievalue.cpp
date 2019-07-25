/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>

#include <xlp/xl/xlpoper.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcellevents/userserievalue.hpp>

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

	template<class Archive>
	void serialize(Archive & ar, NullType& type, const unsigned int version) {
	}

	template<class Archive>
	void serialize(Archive & ar, NAType& type, const unsigned int version) {
	}

	template<class Archive>
	void serialize(Archive & ar, Scalar& scalar, const unsigned int version) {
		ScalarDef sd = scalar.scalar();
		ar & sd;
		scalar = Scalar(sd);
	}

	template<class Archive>
	void serialize(Archive & ar, std::list<Scalar>& values, const unsigned int version) {
		Size size = values.size();
		ar & size;
		values.resize(size);
		std::list<Scalar>::iterator iter = values.begin();
		while (iter != values.end()) { 
			ar & (*iter);
			++iter;
		}
	}
	
	std::string UserSerieValue::idCellValue() const {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 32;
		Xloper xlAddr;
		Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &(*(xlCell_->oper())));
		XlpOper xlpName(&xlAddr, "addr");
		std::string addr = xlpName;
		std::string id = addr + "_" 
			+ boost::lexical_cast<std::string>(xlCell_->row()) + "_"
			+ boost::lexical_cast<std::string>(xlCell_->col());
		return id;
	}

	void UserSerieValue::save(const Scalar& value) const {
		std::list<Scalar> vvalues;
		std::string filename = std::string(XLSession::instance().tmpDir())+"/"+idCellValue()+".xlpar";
		try {
			std::ifstream ifs(filename.c_str());
			if (ifs.good()) {
				boost::archive::text_iarchive ia(ifs);
				ia >> vvalues;
			} else {
				vvalues.resize(size_);
			}
		} catch (...) {
			vvalues.resize(size_);
		}
		if (vvalues.size() > size_) {
			for (Size i=0;i<(vvalues.size()-size_);++i) {
				vvalues.pop_back();
			}
		} else if (vvalues.size() < size_) {
			for (Size i=0;i<(size_-vvalues.size());++i) {
				vvalues.push_back(NullType());
			}
		}
		vvalues.pop_back();
		vvalues.push_front(value);
		std::ofstream ofs(filename.c_str());
    boost::archive::text_oarchive oa(ofs);
		oa << vvalues;
	}

	void UserSerieValue::del() const {
		std::list<Scalar> vvalues;
		std::string filename = std::string(XLSession::instance().tmpDir())+"/"+idCellValue()+".xlpar";
		try {
			std::ifstream ifs(filename.c_str());
			if (ifs.good()) {
				boost::archive::text_iarchive ia(ifs);
				ia >> vvalues;
			} else {
				vvalues.resize(size_);
			}
		} catch (...) {
			vvalues.resize(size_);
		}
		if (vvalues.size() > size_) {
			for (Size i=0;i<(vvalues.size()-size_);++i) {
				vvalues.pop_back();
			}
		} else if (vvalues.size() < size_) {
			for (Size i=0;i<(size_-vvalues.size());++i) {
				vvalues.push_back(NullType());
			}
		}
		vvalues.pop_front();
		vvalues.push_back(NullType());
		std::ofstream ofs(filename.c_str());
    boost::archive::text_oarchive oa(ofs);
		oa << vvalues;
	}

	Scalar UserSerieValue::load(Size i) const {
		try {
			std::list<Scalar> vvalues;
			std::string filename = std::string(XLSession::instance().tmpDir())+"/"+idCellValue()+".xlpar";
			std::ifstream ifs(filename.c_str());
			if (ifs.good()) {
				boost::archive::text_iarchive ia(ifs);
				ia >> vvalues;
				if (i < vvalues.size()) {
					std::list<Scalar>::iterator iter = vvalues.begin();
					for (Size j=0;j<i;++j) {
						++iter;
					}
					return *iter;
				} else {
					return NullType();
				}
			} else {
				return NullType();
			}
		} catch (...) {
			return NullType();
		}
	}

	boost::shared_ptr< Range<Scalar> > UserSerieValue::range() const {
		try {
			std::list<Scalar> vvalues;
			std::string filename = std::string(XLSession::instance().tmpDir())+"/"+idCellValue()+".xlpar";
			std::ifstream ifs(filename.c_str());
			if (ifs.good()) {
				boost::archive::text_iarchive ia(ifs);
				ia >> vvalues;
				boost::shared_ptr< Range<Scalar> > data(new Range<Scalar>(1, vvalues.size()));
				std::copy(vvalues.begin(), vvalues.end(), data->data.begin());
				return data;
			} else {
				return boost::shared_ptr< Range<Scalar> >(new Range<Scalar>(0, 0));
			}
		} catch (...) {
			return boost::shared_ptr< Range<Scalar> >(new Range<Scalar>(0, 0));
		}
	}
	
}

