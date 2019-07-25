/*
 2008 Sebastien Lapedra
*/


#ifndef xlp_singleton_hpp
#define xlp_singleton_hpp

#include <boost/shared_ptr.hpp>
#include <boost/noncopyable.hpp>
#include <map>

namespace xlp {

		template <class T>
		class Singleton {
     public:
       //! access to the unique instance
       static T& instance();
      
		protected:
       Singleton() {}
        /*! derived class might provide an initialization body.
            It will be called after creation
        */
			void initialize() {}

    private:
       Singleton(const Singleton&) {}
    };

    // template definitions

    template <class T>
    T& Singleton<T>::instance() {
        static boost::shared_ptr<T> instance_;
        if (!instance_) {
            instance_ = boost::shared_ptr<T>(new T);
            instance_->initialize();
        }
        return *instance_;
    }
       
}


#endif
