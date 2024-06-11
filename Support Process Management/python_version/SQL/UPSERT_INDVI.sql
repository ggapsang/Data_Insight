INSERT INTO "표준데이터시트_개별속성_240530"

SELECT * FROM "표준데이터시트_centrifugal_pump_수정_20240530" WHERE TRUE

ON CONFLICT("SR_No_ATTR") 
DO UPDATE SET 
	"SRNo" = EXCLUDED."SRNo",
	"속성순번" = EXCLUDED."속성순번",
	"속성값"	= EXCLUDED."속성값",
	"공정" = EXCLUDED."공정",
	"C|C|T" = EXCLUDED."C|C|T",
	"속성명" = EXCLUDED."속성명",
	"카테고리" = EXCLUDED."카테고리",
	"클래스" = EXCLUDED."클래스",
	"타입" = EXCLUDED."타입";
