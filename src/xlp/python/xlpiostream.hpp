/*
 2008 Sebastien Lapedra
*/


#ifndef xlp_iostream_hpp
#define xlp_iostream_hpp

#include <xlp/xlpdefines.hpp>
#include <iosfwd>                      
#include <boost/iostreams/categories.hpp> 
#include <boost/iostreams/stream.hpp>
#include <boost/iostreams/concepts.hpp>
#include <boost/circular_buffer.hpp>
#include <boost/timer.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>



namespace xlp {

	using namespace boost::iostreams;

	class XLPIoStreamDatas {
	public:
	
		struct LineType {
			LineType() : type(" "), msg(" "), index(0) {}
			
			Size index;
			std::string type;
			std::string msg;
		};
		
		XLPIoStreamDatas(Size size);

		void push_back(const std::string& type, const std::string& msg) const;

		Size size();

		boost::circular_buffer<LineType> lines() const;

	private:
		mutable Size index_;
		mutable boost::circular_buffer<LineType> lines_;

	};


	class XLPIoStreamSink : public sink  {
	public:

		XLPIoStreamSink(const std::string& type, const XLPIoStreamDatas& datas);

		std::streamsize write(const char_type* s, std::streamsize n);
		std::string type();

	private:
		std::string type_;
		const XLPIoStreamDatas& datas_;
	};

	}

#endif

