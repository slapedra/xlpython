/*
 2008 Sebastien Lapedra
*/



#ifndef xlp_errors_hpp
#define xlp_errors_hpp

#include <stdlib.h>
#include <iostream>
#include <stdexcept>
#include <string>
#include <boost/current_function.hpp>

namespace xlp {

	/*! error strcuture
   */
	struct ErrorInfo {
		std::string file;
    long line;
    std::string function;
    std::string msg;
  };

  /*! Exception 
   */
  class Error : public std::exception {
		public:
			
			Error(const std::string& file, long line, 
				const std::string& function, const std::string& msg = "");
            
      ~Error() throw() {}
                
      const char* what() const throw ();
                
			ErrorInfo& errorInfo(void);

			const char* message() const throw ();
          
		private:
        
      ErrorInfo errorInfo_;
      std::string lMsg_;
    };

    /*! \def FAIL
    */
    #define FAIL(msg) \
    do { \
        std::ostringstream _os_fail_stream_; \
        _os_fail_stream_ << msg; \
				throw xlp::Error(__FILE__, __LINE__, BOOST_CURRENT_FUNCTION, _os_fail_stream_.str()); \
    } while(false)
    
    /*! \def ASSERT
    */
		#ifdef DEBUG
			#define ASSERT(cond, msg) \
			if (!(cond)) { \
					std::ostringstream _os_assert_stream_; \
					_os_assert_stream_ << msg; \
					throw xlp::Error(__FILE__, __LINE__, BOOST_CURRENT_FUNCTION, _os_assert_stream_.str()); }
		#else
			#define ASSERT(cond, msg)
		#endif
    
    /*! \def REQUIRE
    */
    #define REQUIRE(cond, msg) \
    if (!(cond)) { \
        std::ostringstream _os_require_stream_; \
        _os_require_stream_ << msg; \
				throw xlp::Error(__FILE__, __LINE__, BOOST_CURRENT_FUNCTION, _os_require_stream_.str()); }
    
    /*! \def ENSURE
    */
		#ifdef EXTRA_CHECK
			#define ENSURE(cond, msg) \
			if (!(cond)) { \
				  std::ostringstream _os_ensure_stream_; \
					_os_ensure_stream_ << msg; \
					throw xlp::Error(__FILE__, __LINE__, BOOST_CURRENT_FUNCTION, _os_ensure_stream_.str()); }
		#else	
			#define ENSURE(cond, msg) 
		#endif
}

#endif
