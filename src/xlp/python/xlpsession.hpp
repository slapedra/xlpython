/*
 2008 Sebastien Lapedra
*/


#ifndef xlp_session_hpp
#define xlp_session_hpp

#include <xlp/xlpdefines.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/python/xlpcontainer.hpp>
#include <map>
#include <boost/function.hpp>
#include <boost/bind.hpp>
#include <xlp/python/xlpiostream.hpp>
#include <boost/python/object.hpp>
#include <boost/python/stl_iterator.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>
#include <xlp/util/singleton.hpp>
#include <xlp/xl/xlsession.hpp>
#include <boost/bimap.hpp>
#include <boost/bimap/unordered_set_of.hpp>
#include <boost/bimap/multiset_of.hpp>
#include <boost/multi_index_container.hpp>
#include <boost/multi_index/hashed_index.hpp>
#include <boost/multi_index/member.hpp>
#include <boost/multi_index/ordered_index.hpp>


#include <list>


#define XLP_IOSTREAMBUFFERSIZE	2048
 

namespace xlp {

	#define XLPSESSIONFAIL(msg) \
		do { \
			std::string _pymsg = handleException(); \
			std::ostringstream _xlp_fail_stream_; \
			_xlp_fail_stream_ << msg; \
			if (!_pymsg.empty()) { \
				 _xlp_fail_stream_ << " ( " <<  _pymsg << " )"; \
			} \
			throw Error(__FILE__, __LINE__, __FUNCTION__, _xlp_fail_stream_.str()); \
   } while(false)


	class XLPPyerr {
		public:
			void write(const char* text);
	};

	class XLPPyout {
		public:
			void write(const char* text);
	};

	//!  Excel python session
	/*!
		Main class that defines the different functionalities of XLPython
	*/
	class XLPSession : public Singleton<XLPSession> {
		friend class Singleton<XLPSession>;
	public:
		
		struct CallablePyObjectInfos {
			std::string name;
			bool callable;
			Integer minArgs;
			bool variadic;
			bool instanceMethod;
		};

		 class ReturnHandler {
		 public:
			 ReturnHandler() {
			 }

			 ReturnHandler& operator+=(const python::object& obj) {
				 objs_.append(obj);
				 return *this;
			 }

			 std::string set(const python::object& pyobject, const std::string& output, 
				 bool transpose, bool force, Real wait) const;

			 std::string formula(const python::object& pyobject, const std::string& output, 
				 bool transpose, bool force, Real wait) const;

			 void error(const std::string& msg) const;
			 void warning(const std::string& msg) const;
			 void message(const std::string& msg) const;

			 std::string workbookPath() const;
			 std::string workbookName() const;
			 std::string worksheetName() const;
			 std::string tmpDir() const;
			 Real time() const;

			 void visible(const python::object& sheet, 
				 const python::list& pysheets, bool activate) const;

			 void save(const python::object& pyobject, const std::string& id) const;
			 void load(const std::string& id, const std::string& output, bool transpose) const; 

			 void name(const std::string& scalar, const std::string& name) const;

			 python::object serieValue(const std::string& id, Size index) const;
			 python::object serie(const std::string& id) const;

			 inline std::string simplifiedSet1(const python::object& pyobject, const std::string& outputstr) const {
				 return set(pyobject, outputstr, false, false, 0);
			 }
			 inline std::string simplifiedSet2(const python::object& pyobject, const std::string& outputstr, bool transpose) const {
				 return set(pyobject, outputstr, transpose, false, 0);
			 }
			 inline std::string simplifiedSet3(const python::object& pyobject, const std::string& outputstr, bool transpose, bool force) const {
				 return set(pyobject, outputstr, transpose, force, 0);
			 }

			 inline std::string simplifiedFormula1(const python::object& pyobject, const std::string& outputstr) const {
				 return formula(pyobject, outputstr, false, false, 0);
			 }
			 inline std::string simplifiedFormula2(const python::object& pyobject, const std::string& outputstr, bool transpose) const {
				 return formula(pyobject, outputstr, transpose, false, 0);
			 }
			 inline std::string simplifiedFormula3(const python::object& pyobject, const std::string& outputstr, bool transpose, bool force) const {
				 return formula(pyobject, outputstr, transpose, force, 0);
			 }

			 inline void simplifiedLoad(const std::string& id, const std::string& output) const {
				load(id, output, false);
			 }

			 inline void simplifiedVisible1(const python::object& sheet, const python::list& pysheets) const {
				 visible(sheet, pysheets, false);
			 }

			 inline void simplifiedVisible2(const python::object& sheet) const {
				 python::list l;
				 visible(sheet, l, false);
			 }

			 const python::list& objects() {
				 return objs_;
			 }

			private:
				python::list objs_;
				XLCaller caller_;
			};

			XLPSession();

			void init() {}

			void initialize();

			~XLPSession();

			boost::iostreams::stream<XLPIoStreamSink>& cerr() { return xlpCerr_; }
			boost::iostreams::stream<XLPIoStreamSink>& cout() { return xlpCout_; }
			boost::iostreams::stream<XLPIoStreamSink>& pyerr() { return xlpPyerr_; }
			boost::iostreams::stream<XLPIoStreamSink>& pyout() { return xlpPyout_; }
			XLPIoStreamDatas& ioStreamDatas() { return xlpIoStreamDatas_; }

			std::string type(const Scalar& scalar) const;

			std::string pythonPath();
			std::string pythonVersion();
			std::string workbookPath();
			std::string workbookName();
			std::string worksheetName(bool little = true);

			boost::shared_ptr<python::dict> localDict(
				const boost::shared_ptr< Range<Scalar> >& args,
				const boost::shared_ptr< Range<Scalar> >& values, 
	  			const boost::shared_ptr< Range<Scalar> >& index) const;
		 
		 std::string set(const Scalar& scalar, 
			const boost::shared_ptr< Range<Scalar> >& args,
			const boost::shared_ptr< Range<Scalar> >& values, 
	  		const boost::shared_ptr< Range<Scalar> >& index,
 			const std::string& id) const;

		 boost::shared_ptr< Range<Scalar> > get(const Scalar& scalar, 
			const boost::shared_ptr< Range<Scalar> >& args, bool force = false) const;

		 std::string apply(const Scalar& scalar,
			const boost::shared_ptr< Range<Scalar> >& args, const std::string& id) const;
		 boost::shared_ptr< Range<Scalar> > getApply(const Scalar& scalar,
			const boost::shared_ptr< Range<Scalar> >& args, bool force = false) const;

		 std::string attr(const Scalar& scalar, const std::string& attr,
			const boost::shared_ptr< Range<Scalar> >& args, const std::string& id) const;
		 boost::shared_ptr< Range<Scalar> > getAttr(const Scalar& scalar, const std::string& attr,
			const boost::shared_ptr< Range<Scalar> >& args, bool force = false) const;

		 boost::shared_ptr< Range<Scalar> > listAllObjects() const;
		 void deleteAllObjects() const;
		 void deleteObject(const Scalar& scalar) const;
		 boost::shared_ptr< Range<Scalar> > listAllAttr(const Scalar& scalar) const;
		 boost::shared_ptr< Range<Scalar> > attrInfos(const Scalar& scalar, const std::string& attr) const;
		 
		 std::string loadFile(const std::string& filename, const std::string& id) const;
		 
		 std::string import(const std::string& module, const std::string& id) const;
		 void fromImport(const std::string& module) const;
		 
		 boost::shared_ptr< Range<Scalar> > exec(
			const boost::shared_ptr< Range<Scalar> >& scalarCode, 
			const boost::shared_ptr< Range<Scalar> >& args, 
			const boost::shared_ptr< Range<Scalar> >& ids) const;

		 void compile(
			const boost::shared_ptr< Range<Scalar> >& scalarCode, const std::string& id) const;

		 void insertPaths(const std::vector<std::string>& paths) const;
		 boost::shared_ptr< Range<Scalar> > XLPSession::sysPath() const;

		 std::string workbook(const std::string& id) const;

		 void onWorkbookClose(const std::string& name) const;

		 python::object scalarToPyObj(const Scalar& scalar) const;


		 boost::shared_ptr< Range<Scalar> > excelId(const boost::shared_ptr< Range<Scalar> >& range) const;
		 boost::shared_ptr< Range<Scalar> > pythonId(const boost::shared_ptr< Range<Scalar> >& range) const;

		private:

			mutable python::object module_;
			mutable python::dict session_;

			mutable bool returnHandlerInit_;
 
			typedef std::map<std::string, boost::shared_ptr<XLPContainerOutBindingMethod> > OutMap;
			OutMap containerOutBinds_;
			mutable boost::shared_ptr< std::list<std::string> > tmpObjectIds_;

			struct Code {
				boost::python::object code;
				std::string workbook;
				std::string id;

				Code(const boost::python::object& scode, const std::string& sid, const std::string& sworkbook)
				: code(scode), id(sid), workbook(sworkbook)  {
				}
			};

			struct code_key:composite_key<
				Code,
				BOOST_MULTI_INDEX_MEMBER(Code, std::string, id),
				BOOST_MULTI_INDEX_MEMBER(Code, std::string, workbook)
			>{};

			typedef multi_index_container<
				Code,
				indexed_by<
					ordered_unique<code_key>
				>
			> CodeMap;

			mutable CodeMap codeMap_;
		
			XLPIoStreamDatas xlpIoStreamDatas_;
			boost::iostreams::stream<XLPIoStreamSink> xlpCout_;
			boost::iostreams::stream<XLPIoStreamSink> xlpCerr_;
			boost::iostreams::stream<XLPIoStreamSink> xlpPyout_;
			boost::iostreams::stream<XLPIoStreamSink> xlpPyerr_;
			std::streambuf* oldCout_;
			std::streambuf* oldCerr_;

			typedef boost::bimap<
				boost::bimaps::unordered_set_of<std::string>,
				boost::bimaps::multiset_of<std::string, std::less<std::string>>
			> PyObjectBimap;

			typedef PyObjectBimap::value_type PyObjectBimapElement;

			mutable PyObjectBimap workbookObjects_;

			void insertObject(const python::object& obj, 
				const std::string& id, const std::string& workbook) const;

			void validateObjectId(const std::string& id) const;
			std::string createObjectId(const python::object& obj, const std::string& id) const;
			std::string createExcelId(const std::string& id) const;
			std::string getType(const python::object& obj) const;
			Scalar convertPyObjectToScalar(const python::object& obj, bool force = false) const;
			python::object forceConversion(const python::object& obj) const;

			std::string handleException() const;
			bool deleteObject(const std::string& id) const;

			boost::shared_ptr< Range<Scalar> > get(const python::object& obj, bool force = false) const;
		
			python::object applyScalar(const Scalar& scalar, const boost::shared_ptr< Range<Scalar> >& args) const;
		
			python::object attrScalar(const Scalar& scalar, const std::string& attr,
				const boost::shared_ptr< Range<Scalar> >& args) const;

			// insert data in an object
			void insert(const python::object& obj, const python::dict& local) const;

			CallablePyObjectInfos getCallableObjectInfos(const python::object& obj) const;

			std::string searchName(const python::object& obj) const;

			std::vector<std::string> XLPSession::getSysPath() const;;
		};

	

}

#endif

