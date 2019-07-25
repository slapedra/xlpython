/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/python/xlpiostream.hpp>

namespace xlp {

	XLPIoStreamDatas::XLPIoStreamDatas(Size size) 
	: lines_(size), index_(0) {
	}

	void XLPIoStreamDatas::push_back(const std::string& type, const std::string& msg) const {
		XLPIoStreamDatas::LineType line;
		line.type = type;
		line.msg = msg;
		line.index = index_;
		lines_.push_back(line);
		index_++;
	}

	Size XLPIoStreamDatas::size() {
		return lines_.max_size();
	}

	boost::circular_buffer<XLPIoStreamDatas::LineType> XLPIoStreamDatas::lines() const {
		return lines_;
	}

	XLPIoStreamSink::XLPIoStreamSink(const std::string& type, const XLPIoStreamDatas& datas) 
	: datas_(datas), type_(type) { 
	}

	std::streamsize XLPIoStreamSink::write(const char_type* s, std::streamsize n) {
		static std::string tmp;
		char *t = (char *)s;
		for (Size i=0;i<Size(n);i++) {
			if (*t ==  '\n') {
				datas_.push_back(type_, tmp);
				tmp = "";
			} else {
				tmp += *t;
				t++;
			}
		}
		return n;
	}

	std::string XLPIoStreamSink::type() {
		return type_;
	}
}



