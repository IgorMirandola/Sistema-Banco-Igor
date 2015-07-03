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
-- Table structure for table `line`
--

DROP TABLE IF EXISTS `line`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `line` (
  `lineID` int(11) NOT NULL AUTO_INCREMENT,
  `caseID` int(11) NOT NULL,
  `inicialBusNumber` int(10) unsigned NOT NULL,
  `finalBusNumber` int(10) unsigned NOT NULL,
  `sequencialNumber` int(11) DEFAULT NULL,
  `length` double DEFAULT NULL,
  `resistence` double DEFAULT NULL,
  `reactance` double DEFAULT NULL,
  `shuntSusceptance` double DEFAULT NULL,
  `rating1` double DEFAULT NULL,
  `rating2` double DEFAULT NULL,
  `rating3` double DEFAULT NULL,
  `description` varchar(45) DEFAULT NULL,
  `circuitoNumber` int(11) DEFAULT NULL,
  PRIMARY KEY (`lineID`,`inicialBusNumber`,`caseID`,`finalBusNumber`),
  UNIQUE KEY `LineID_UNIQUE` (`lineID`),
  KEY `inicialBusNumber_idx` (`caseID`,`inicialBusNumber`),
  KEY `finalBusNumber_idx` (`finalBusNumber`,`caseID`),
  CONSTRAINT `finalBusNumber` FOREIGN KEY (`finalBusNumber`, `caseID`) REFERENCES `bus` (`Bus Number`, `case ID`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `inicialBusNumber` FOREIGN KEY (`caseID`, `inicialBusNumber`) REFERENCES `bus` (`case ID`, `Bus Number`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `line`
--

LOCK TABLES `line` WRITE;
/*!40000 ALTER TABLE `line` DISABLE KEYS */;
/*!40000 ALTER TABLE `line` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2015-07-02 22:44:17
