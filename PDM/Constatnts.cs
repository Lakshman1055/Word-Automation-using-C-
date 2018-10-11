using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDM
{
    public class Constatnts
    {
        public const string statement1 = @"SELECT SN.name Synonym_NAME,Replace (Replace (Replace (SN.base_object_name, '[', ''), ']', ''),'dbo.','')AS Table_NAME FROM Sys.synonyms SN WHERE base_object_name LIKE '%govt%'";

        public const string statement2 = @"SELECT TC.COLUMN_NAME,
       TC.ORDINAL_POSITION,
       TC.COLUMN_DEFAULT,
       TC.IS_NULLABLE,
       UPPER (TC.DATA_TYPE) DATA_TYPE,
       TC.CHARACTER_MAXIMUM_LENGTH [LENGTH],
       TC.NUMERIC_PRECISION [PRECISION],
       TC.NUMERIC_SCALE [SCALE],
          COLUMNPROPERTY(object_id(TC.TABLE_NAME), TC.COLUMN_NAME, 'IsIdentity')  IsIdentity ,
          IDENT_SEED(T.TABLE_NAME) AS Seed,
          IDENT_INCR(T.TABLE_NAME) AS Increment,
          IDENT_CURRENT(T.TABLE_NAME) AS Current_Identity
  FROM INFORMATION_SCHEMA.COLUMNS TC 
    JOIN INFORMATION_SCHEMA.TABLES T 
         ON TC.TABLE_SCHEMA = T.TABLE_SCHEMA 
           AND TC.TABLE_NAME = T.TABLE_NAME 
 WHERE TC.TABLE_SCHEMA = 'dbo'
   AND TC.TABLE_NAME = @As_Table_NAME;";

        public const string statement3 = @"SELECT C.TABLE_SCHEMA,C.TABLE_NAME ,C.CONSTRAINT_NAME, CC.COLUMN_NAME
  FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS C
    JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE CC 
         ON C.CONSTRAINT_NAME = CC.Constraint_name 
          AND C.TABLE_SCHEMA = CC.TABLE_SCHEMA
          AND C.TABLE_NAME  = CC.TABLE_NAME
  WHERE C.TABLE_SCHEMA = 'dbo'
   AND  C.TABLE_NAME = @As_Table_NAME
   AND  C.CONSTRAINT_TYPE = 'Primary Key';";
    }

    
}
