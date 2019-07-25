/*
 2008 Sebastien Lapedra
*/

#ifndef xlcommandevent_hpp
#define xlcommandevent_hpp

namespace xlp {
 
	class XLCommandEvent {
	public:
		virtual bool execute(boost::shared_ptr<XLRef>& xlCell) const = 0;
	};
	
}

#endif
