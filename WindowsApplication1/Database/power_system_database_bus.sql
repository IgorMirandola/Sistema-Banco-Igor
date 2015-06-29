-- MySQL dump 10.13  Distrib 5.6.23, for Win64 (x86_64)
--
-- Host: localhost    Database: power_system_database
-- ------------------------------------------------------
-- Server version	5.6.25-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `bus`
--

DROP TABLE IF EXISTS `bus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `bus` (
  `Bus Number` int(10) unsigned NOT NULL,
  `case ID` int(11) NOT NULL,
  `Sequencial Number` int(11) DEFAULT NULL,
  `Bus name` varchar(255) DEFAULT NULL,
  `Voltage` double DEFAULT NULL,
  `Phase` double DEFAULT NULL,
  `Voltage Base` double DEFAULT NULL,
  `Desired Voltage` double DEFAULT NULL,
  `Max Power Voltage` double DEFAULT NULL,
  `Min Power Voltage` double DEFAULT NULL,
  PRIMARY KEY (`Bus Number`,`case ID`),
  KEY `case ID_idx` (`case ID`),
  CONSTRAINT `case ID` FOREIGN KEY (`case ID`) REFERENCES `power_system_case` (`ID`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='This the table for buses. ';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `bus`
--

LOCK TABLES `bus` WRITE;
/*!40000 ALTER TABLE `bus` DISABLE KEYS */;
INSERT INTO `bus` VALUES (1,14,1,'Bus 1',1.06,0,0,1.06,0,0),(2,14,2,'Bus 2',1.045,-4.98,0,1.045,50,-40),(3,14,3,'Bus 3',1.01,-12.72,0,1.01,40,0),(4,14,4,'Bus 4',1.019,-10.33,0,0,0,0),(5,14,5,'Bus 5',1.02,-8.78,0,0,0,0),(6,14,6,'Bus 6',1.07,-14.22,0,1.07,24,-6),(7,14,7,'Bus 7',1.062,-13.37,0,0,0,0),(8,14,8,'Bus 8',1.09,-13.36,0,1.09,24,-6),(9,14,9,'Bus 9',1.056,-14.94,0,0,0,0),(10,14,10,'Bus 10',1.051,-15.1,0,0,0,0),(11,14,11,'Bus 11',1.057,-14.79,0,0,0,0),(12,14,12,'Bus 12',1.055,-15.07,0,0,0,0),(13,14,13,'Bus 13',1.05,-15.16,0,0,0,0),(14,14,14,'Bus 14',1.036,-16.04,0,0,0,0),(611,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(632,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(633,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(634,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(645,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(646,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(650,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(652,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(671,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(675,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(680,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(684,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(692,29,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);
/*!40000 ALTER TABLE `bus` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2015-06-29  2:27:31
