require 'rubygems'
require 'bundler'

require 'java'

Bundler.require
require './document_creator'
Dir["./lib/*.rb"].each do |rb|
  require rb
end

DocumentCreator.run!
