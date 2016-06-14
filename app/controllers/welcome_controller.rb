class WelcomeController < ApplicationController
  def index
    @file_names = Xlsx.file_names
  end
end

