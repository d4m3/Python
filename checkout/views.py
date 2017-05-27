from django.conf import settings
from django.contrib.auth.decorators import login_required
from django.shortcuts import render
import stripe

#stripe.api_key = settings.STRIPE_SECRET_KEY

@login_required
def checkout(request):
	if request.method == 'POST':
		token = request.POST#['stripeToken']
		print(token)
	context = {}
	template = 'checkout.html'
	return render(request, template, context)

"""
# Create your views here.
@login_required
def checkout(request):
	publishKey = settings.STRIPE_PUBLISHABLE_KEY
	if request.method == 'POST':
		token = request.POST['stripeToken']
		#token = request.POST['csrfmiddlewaretoken']
		# Token is created using Stripe.js or Checkout!
		# Get the payment token submitted by the form:
		# Charge the user's card:
		try:
			charge = stripe.Charge.create(
		  		amount=1000, # Amount in cents
		  		currency="usd",
		  		source=token,
		  		description="Example charge"
		  		#metadata={"order_id": 6735},
			)
		except stripe.error.CardError as e:
			# The card has been declined
			pass
	context = {'publishKey': publishKey}
	template = 'checkout.html'
	return render(request, template, context)
"""