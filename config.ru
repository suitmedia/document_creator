require 'rubygems'
require 'bundler'

include Java

Bundler.require
require './document_creator'
Dir["./lib/*.rb"].each do |rb|
  require rb
end

DocumentCreator.run!
