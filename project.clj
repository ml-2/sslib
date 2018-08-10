(defproject sslib "0.2.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "GNU General Public License version 3"
            :url ""} ;; TODO
  :dependencies [[org.clojure/clojure "1.9.0"]
                 [dataorigin/dataorigin "0.1.0-SNAPSHOT"]
                 [org.apache.poi/poi "3.17"]
                 [org.apache.poi/poi-ooxml "3.17"]]
  :aot [sslib.core])
