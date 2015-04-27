# xlsx-io
XLSX Input/Output library

#Oracle support
To support oracle database, you need to manually download ojdbc JAR from oracle site (http://www.oracle.com/technology/software/tech/java/sqlj_jdbc/index.html)
After downloading, you need to install the JAR using the following command:
mvn install:install-file -DgroupId=ojdbc -DartifactId=ojdbc -Dversion=14 -Dpackaging=jar -Dfile=ojdbc7.jar

