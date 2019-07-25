/*
 2008 Sebastien Lapedra
*/


#ifndef xlsession_hpp
#define xlsession_hpp

#define NOMINMAX
#include <windows.h>
#include <xlp/xl/xldefines.hpp>
#include <xlp/xl/xlcall.h>
#include <xlp/xl/framewrk.hpp>
#include <xlp/xl/xlpoper.hpp>
#include <xlp/xl/xlref.hpp>
#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/util/singleton.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>
#include <sstream>
#include <stdlib.h> 
#include <xlp/xl/xloper.hpp>
#include <list>
#include <set>
#include <boost/timer.hpp>
#include <boost/function.hpp>
#include <boost/multi_index_container.hpp>
#include <boost/multi_index/ordered_index.hpp>
#include <boost/multi_index/identity.hpp>
#include <boost/multi_index/member.hpp>
#include <boost/multi_index/composite_key.hpp>
#include <boost/tuple/tuple.hpp>

#define			XL_MIN_KEY_TIME					0.1

#define			XL_ACTION_WAITTIME			0.5 / (24.0 * 60.0 * 60.0)

using boost::multi_index_container;
using namespace boost::multi_index;

namespace xlp {

	/*! XLCurrentPosition
	 */
	class XLCurrentPosition {
	public:		 
		XLCurrentPosition(long id, Size row, Size col, const std::string& addr,
			const std::string& workbook, bool transpose);

		std::string idSheet() const;
		std::string idPosition() const;
			 
		inline const std::string& addr() const { return addr_; }
		inline const std::string& workbook() const { return workbook_; }

		void addRow() const;
		void addCol() const;

	 private:
		bool transpose_;
		mutable Size curRow_;
		mutable Size curCol_;
		Size curIdSheet_;
		std::string addr_;
		std::string workbook_;

		std::string computeNumberRepresentation(long id) const;
	};

	//!  Excel Caller infos 
	/*!
		Excel Caller infos
	*/
	class XLCaller {
	public:
		XLCaller ();

		inline Size row() const { return row_; }
		inline Size rows() const { return rows_; }
		inline Size col() const { return col_; }
		inline Size cols() const { return cols_; }
		inline const std::string& addr() const { return addr_; }
		inline const std::string& workbook() const { return workbook_; }
		inline long idSheet() const { return idSheet_; }

	private:
		Size row_;
		Size col_;
		Size rows_;
		Size cols_;
		long idSheet_;
		std::string addr_;
		std::string workbook_;
	};

	//!  Excel function class for registration
	/*!
		Excel function class for registration
	*/
	class XLFunction {
	public:

		XLFunction(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);

		~XLFunction() {}

		inline std::string cName() const {return cName_;} 

		inline std::string xlName() const {return xlName_;} 
		
		inline std::string cArgs() const {return cArgs_;} 

		inline std::string xlArgs() const {return xlArgs_;} 
	
		inline std::string desc() const {return desc_;} 

		inline std::string module() const {return moduleName_;}

		inline std::vector<std::string> descArgs() const {return descArgs_;} 

		void registerFunction();

		void unregisterFunction();

	private:
		double regId_;
		std::string moduleName_;
		std::string cName_;
		std::string xlName_;
		std::string cArgs_;
		std::string xlArgs_;
		std::string desc_;
		std::vector<std::string> descArgs_;
						
	};

	
	//!  Excel command class for registration
	/*!
		Excel command class for registration
	*/
	class XLCommand {
	public:

		XLCommand(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);

		~XLCommand() {}

		inline std::string cName() const {return cName_;} 

		inline std::string xlName() const {return xlName_;} 
		
		inline std::string cArgs() const {return cArgs_;} 

		inline std::string xlArgs() const {return xlArgs_;} 
	
		inline std::string desc() const {return desc_;} 

		inline std::string module() const {return moduleName_;}

		inline std::vector<std::string> descArgs() const {return descArgs_;} 

		void registerCommand();

		void unregisterCommand();

	private:
		double regId_;
		std::string moduleName_;
		std::string cName_;
		std::string xlName_;
		std::string cArgs_;
		std::string xlArgs_;
		std::string desc_;
		std::vector<std::string> descArgs_;
						
	};

	//!  Excel session 
	/*!
		Excel session that store all the functions for registration 
		and unregistration procedure
	*/
	class XLSession : public Singleton<XLSession> {
	friend class Singleton<XLSession>;
	public:
		XLSession()
		: runningRealtime_(false), timestep_(XL_TIMESTEP), 
			output_(""), transpose_(false), onWorkbookCloseCallback_(NULL), 
			nameKey_("xl"), activate_(true), developpermode_(false) {
		}

		~XLSession();

		void validateTrigger(OPER* xlTrigger);
		bool convertBool(OPER* xlValue, bool defaultValue);

		void init() {}
		void initialize() {}

		void registerSession();
		void unregisterSession();

		std::string id() const { return id_; }
		std::string workbookPath() const;
		std::string workbookName() const;
		std::string worksheetName(bool little = true) const;

		void push_back(const boost::shared_ptr<XLFunction>& func);
		void push_back(const boost::shared_ptr<XLCommand>& command);

		inline bool runningRealtime() const { return runningRealtime_; }

		void activate(bool flag) const { activate_ = flag; }
		inline bool activated() const { return activate_; }

		void onDevelopement(bool flag) const { developpermode_ = flag; }
		inline bool developperMode() const { return developpermode_; }

		void calculateNow() const;

		Real time() const;
		inline Real timestep() const { return timestep_; }; 
		void realtime(bool flag, Real step = XL_TIMESTEP) const;

		Scalar serieValue(const boost::shared_ptr<XLRef>& xlRef, Size index) const;
		boost::shared_ptr< Range<Scalar> > serie(const boost::shared_ptr<XLRef>& xlRef) const;

		boost::shared_ptr<XLRef> checkOutput(const std::string& output) const;
		std::string fillOutput(const boost::shared_ptr< Range<Scalar> >& datas) const;
		std::string fillOutput(const boost::shared_ptr<XLRef>& ref, 
			const boost::shared_ptr< Range<Scalar> >& datas, bool transpose, Real time = 0) const;
		std::string formula(const boost::shared_ptr< Range<Scalar> >& data, bool transpose, 
			const std::string& formula, Real time = 0) const;
		std::string formula(const boost::shared_ptr<XLRef>& ref,
			const boost::shared_ptr< Range<Scalar> >& data, bool transpose, Real time = 0) const;
		void defineName(const std::string& scalar, const std::string& name) const;
		void defineName(const boost::shared_ptr<XLRef>& xlRef, const std::string& name) const;

		void onRecalc() const;
		void onTime() const;
		void onSheet() const;
		void onOpen() const;
		void onWindow() const;
		void onEntry() const;
		void onDC() const;
		void onDel() const;
		void onSKey() const;
		void onHideProtect(bool flag = true) const;

		inline void setOnWorkbookClose(void (*callback)(const std::string&)) const { onWorkbookCloseCallback_ = callback; }
		inline void setNameKey(const std::string& key) const { nameKey_ = key; }
		inline std::string nameKey() const { return nameKey_; }

		void freeze(bool flag, bool display = true) const;

		void unprotect(bool flag) const;

		void visible(const std::vector<std::string>& sheets, const std::string& sheet, bool activate) const;

		bool isProtected(const std::string& sheet = "") const;

		boost::shared_ptr<XLCurrentPosition> currentPosition() const;
		boost::shared_ptr<XLRef> currentRef(const XLCaller& caller, const std::string& output) const;	

		inline bool transpose() { return transpose_; }

		void enterFunction(OPER* xlTrigger, OPER* xlTranspose = NULL, OPER* xlOutput = NULL) const;
		void exitFunction() const;

		void message(const std::string& str) const;
		void error(const std::string& str) const;
		void warning(const std::string& str) const;

		std::string tmpDir() const;
		void saveRange(const boost::shared_ptr< Range<Scalar> >& range, const std::string& id) const;
		boost::shared_ptr< Range<Scalar> > loadRange(const std::string& id) const;
		
		boost::shared_ptr<XLCellEventMap> cellEvents(const std::string& src, 
			const std::list< boost::shared_ptr<XLRef> >& xlRefs, bool flag = false) const;

		inline boost::shared_ptr<XLCellEventMap> cellEvents(const std::string& src, 
			const boost::shared_ptr<XLRef>& xlRef) const {
			std::list< boost::shared_ptr<XLRef> > xlRefs;
			xlRefs.push_back(xlRef);
			return cellEvents(src, xlRefs); 
		}

	private:
		mutable bool activate_;
		mutable bool developpermode_;
		mutable std::string worksheet_;
		mutable std::map<std::string, boost::shared_ptr<XLCellEventMap> > cellmaps_;

		mutable boost::timer timer_;
		double firstTime_;
		std::string id_;
		std::vector< boost::shared_ptr<XLFunction> > functions_;
		std::vector< boost::shared_ptr<XLCommand> > commands_;
		mutable void (*onWorkbookCloseCallback_)(const std::string&);

		mutable std::list< boost::shared_ptr<XLCommandEvent> > recalcEvents_;
		mutable bool runningRealtime_;
		mutable Real timestep_;
		mutable bool transpose_;
		mutable std::string output_;
		mutable bool flagCurrentPosition_;
		mutable boost::shared_ptr<XLCurrentPosition> currentPosition_;
		mutable boost::shared_ptr<XLRef> xlref_;
		mutable std::set<std::string> workbooks_;
		mutable std::string nameKey_;
	};

	extern boost::shared_ptr<XLFunction> function(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);

	extern boost::shared_ptr<XLCommand> command(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);

	extern void registerXLFunction(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);

	extern void registerXLCommand(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs);
	
}

#endif
