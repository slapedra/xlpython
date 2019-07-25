/*
 2008 Sebastien Lapedra
*/

#ifndef scalaropervisitor_hpp
#define scalaropervisitor_hpp

#include <xlp/xl/xldefines.hpp>
#include <xlp/xl/xlsession.hpp>
#include <boost/variant.hpp>

using namespace boost;

namespace xlp {

	class ScalarToOperVisitor : public boost::static_visitor<> {
	public:
		ScalarToOperVisitor(OPER *oper, bool dllfree = true) :
		oper_(oper), dllfree_(dllfree) {
		}

		void operator()(const Integer &i) const;
		void operator()(const Real &r) const;
    void operator()(const bool &b) const;
		void operator()(const std::string&) const;
		void operator()(const NullType&) const;
		void operator()(const NAType&) const;
	private:
		OPER *oper_;
		bool dllfree_;
	};
	
}


#endif

