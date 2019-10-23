(defproject sslib "0.5.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "GNU General Public License version 3"
            :url ""} ;; TODO
  :dependencies [[org.clojure/clojure "1.9.0"]
                 [dataorigin/dataorigin "0.1.0-SNAPSHOT"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]]
  :aot [sslib.core])
