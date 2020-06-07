(defproject sslib "0.5.0-SNAPSHOT"
  :description "Library for interacting with xls and xlsx files via Apache POI"
  :url ""
  :license {:name "GNU General Public License version 3 or later, with Clojure exception"
            :url ""}
  :dependencies [[org.clojure/clojure "1.9.0"]
                 [dataorigin/dataorigin "0.1.0-SNAPSHOT"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]]
  :aot [sslib.core])
