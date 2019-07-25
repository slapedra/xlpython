/*
 2008 Sebastien Lapedra
*/


#include <sstream>
#include <stdexcept>
#include <xlp/util/errors.hpp>

namespace xlp {

	Error::Error(const std::string& file, long line,
		const std::string& function, const std::string& msg) {
		std::ostringstream omsg;
		std::string::size_type n = file.find_last_of("/\\");
        
		errorInfo_.file = file;
		errorInfo_.line = line;
		errorInfo_.function = function; 
		errorInfo_.msg = msg;
		if (n != std::string::npos) 
			errorInfo_.file = errorInfo_.file.substr(n+1);
		omsg << "(" << errorInfo_.file << ", " << errorInfo_.line  
			<< ", " << errorInfo_.function \
			<< ") : " << errorInfo_.msg; 
		lMsg_ = omsg.str();
	}
    
  const char* Error::what() const throw () {
		//return lMsg_.c_str();
		return errorInfo_.msg.c_str();
  }
  
	ErrorInfo& Error::errorInfo(void) {
		return errorInfo_;
  }

	const char* Error::message() const throw () {
		return errorInfo_.msg.c_str();
  }

}

