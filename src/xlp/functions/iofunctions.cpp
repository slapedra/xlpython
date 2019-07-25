/*
 2008 Sebastien Lapedra

*/

#include <xlp/config.hpp>
#include <limits>
#include <algorithm>
#include <xlp/xl/xl.hpp>
#include <xlp/xlpdefines.hpp>
#include <xlp/python/xlpsession.hpp>
#include <boost/timer.hpp>


namespace xlp {

	
	XL_EXPORT OPER *xlpFullIO(OPER *xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XLPIoStreamDatas& datas = XLPSession::instance().ioStreamDatas();
		boost::circular_buffer<XLPIoStreamDatas::LineType> lines = datas.lines();
		boost::shared_ptr< Range<Scalar> > scalars(new Range<Scalar>(lines.size(), 3) );
		Size offset = 0;
		for (Size i=0;i<lines.size();i++) {
			scalars->data[offset] = Integer(lines[i].index);
			offset++;
			scalars->data[offset] = lines[i].type;
			offset++;
			scalars->data[offset] = lines[i].msg;
			offset++;
		}
		XL_END(scalars)
	}

	XL_EXPORT OPER *xlpIO(OPER* xlLines, OPER *xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpLines(xlLines, "lines", true, false);
		Integer nblines = xlpLines;
		REQUIRE (nblines >0, "bad number of lines");
		XLPIoStreamDatas& datas = XLPSession::instance().ioStreamDatas();
		boost::circular_buffer<XLPIoStreamDatas::LineType> lines = datas.lines();
		boost::shared_ptr< Range<Scalar> > scalars;
		if (lines.size()<= 0) { 
			scalars.reset(new Range<Scalar>(1, 1) );
			scalars->data[0] = "IO empty";
		} else {
			if (Size(nblines) > Size(lines.size())) {
				nblines = lines.size();
			}
			scalars.reset(new Range<Scalar>(nblines+1, 3) );
			Size offset = 0;
			scalars->data[offset]= "last";
			offset++;
			scalars->data[offset] = lines.back().type;
			offset++;
			scalars->data[offset] = lines.back().msg;
			offset++;
			for (int i=nblines-1;i>=0;i--) {
				scalars->data[offset] = Integer(lines[lines.size()-i-1].index);
				offset++;
				scalars->data[offset] = lines[lines.size()-i-1].type;
				offset++;
				scalars->data[offset] = lines[lines.size()-i-1].msg;
				offset++;
			}
		}
		XL_END(scalars)
	}

}



