 //****************************************************************************************************
 //
 //     3      Vad gör funktionen?
 //
 //**************************************************************************************************** 


 /*
  * Note: Majority of this evaluation is dependent on 6 seemingly asynchronous elements are loaded,
  * otherwise there are multiple loading-bars rendered which will most likely be updated once the data is loaded.
  * 
  * The entire function revolves around groups of bookings.
  * A group is extracted from a selected unit.
  * I did not quite understand what these groups, or units could represent (Restaurants?, Companies?, Set of tables?),
  * But the groups have a relation to the individual articles, so one theory could be that a group is
  * a filter of all bookings that include a specific article-type.
  * 
  * Each group gets a list of bookings, and each booking acquires a list of articles.
  * The function returns a column layout of two columns, which will be inserted and rendered by the caller.
  * 
  * Column 1:  Unit Controller
  * A table of "units", when selecting a unit the "handleSelectUnit" will be called,
  * this will probably re-run this function based on a new selected unit.
  * 
  * Column 2:  Group-render.
  * A List of groups, which each contains a list of bookings, with their individual list of articles.
  * The content of these object is what the functions main evaluation is about
  * (Seem to not be dictated how these are ordered in this function other then they are contained in a VerticalLayout, 
  * suggesting they will produce a vertical scrollwheel downwards if overflown from the window-size)
  * 
  * This Group-render is controlled by a "dateControl" which handle a time period, suggesting you can filter the rendering based on time period.
  * 
  * Below are comments or explanations of the evaluation.
  */

 function explainThisFunc() {
	var self = this
	var groupsToRender = []
	var hasEditPermissions = this.hasPermission('booking.update')
	var actions = []

    // Checking for edit permission and adds a "new booking button"
	if (hasEditPermissions) {
		actions.push(
			<Button
				id="addBooking"
				preset="addPerson"
				onClick={this.handleNewBookingClick}
				disabled={!hasEditPermissions}
			/>,
		)
	}

    // Checking if asynchronous call is finished and accessible
	var allLoaded =
		isLoaded(this.state.bookingsResult) &&
		isLoaded(this.state.bookedArticlesResult) &&
		isLoaded(this.state.currentUserResult) &&
		isLoaded(this.state.timePeriodStatisticsResult) &&
		isLoaded(this.state.tablesResult) &&
		isLoaded(this.state.tableArticlesResult)

	if (this.state.selectedUnit && allLoaded) {
		var tables = this.state.tablesResult.value

		var groups = this.state.selectedUnit.articleGroups

        // Filter invalid Groups
		var filteredGroups = groups.filter(
			(articleGroup) =>
				tables.findIndex((t) => t.articleGroupId === articleGroup.id) > -1,
		)

		var unspecifiedGroup = ArticleGroup.Record({
			articleGroupId: '0',
			name: this.t('booking.hasNoTables'),
		})

        // Add an unspecified group
		filteredGroups = filteredGroups.push(unspecifiedGroup)


        // For each group of filteredGroup
		filteredGroups.forEach(function (group) {
			var bookings = Immutable.List()

            // For each booking from state.bookingsResult
			self.state.bookingsResult.value.forEach(function (booking) {
                // Add all articles associated with the current booking
				booking = booking.set(
					'articles',
					BookingStore.getBookedArticlesForBookingResult(
						booking.id,
						self.state.date,
						self.state.selectedUnit.id,
					).value,
				)

                // Constructs a filtered list of the current bookings' articles,
                // filter based on the current group
				var articlesInGroup = booking.articles.filter(
					(a) => a.articleGroupId === group.id,
				)

                // If there are any articles in the current booking, associated by the current group
				// (or if there are no articles and its the unspecified group)
				//
                // Add this booking (with all its articles) to the list of bookings
                // that will be related to the current group
				if (
					articlesInGroup.count() > 0 ||
					(booking.articles.count() < 1 && group === unspecifiedGroup)
				) {
					bookings = bookings.push(booking)
				}
            }) 
            // End of foreach booking in group.
            // At this point, "bookings" contain all bookings related to the current group.

            // Sort the bookings based on start-time
            bookings = bookings.sort(function (a, b) {
				var moment1 = Moment(a.start, 'YYYY-MM-DDTHH:mm')

				var moment2 = Moment(b.start, 'YYYY-MM-DDTHH:mm')

				return moment1.diff(moment2, 'minutes')
			})


            // save the current group as a <BookingGroup> with all its bookings as a list of <BookingListItem>
			var bookingsToRender = []
			bookings.forEach(function (booking) {
				bookingsToRender.push(
					<BookingListItem
						booking={booking}
						editBooking={self.state.editBooking}
						contacts={self.state.contacts}
						onChange={self.onChange}
						onClick={self.onClick}
					/>,
				)
			})

			groupsToRender.push(
				<BookingGroup
					key={group.name}
					title={group.name}
					object={group}
					editBooking={self.state.editBooking}
					contacts={self.state.contacts}
					expanded={false}
				>
					{bookingsToRender}
				</BookingGroup>,
			)
        })
        //End of foreach group
	}


	var hasValidDaynote = this.hasValidDaynote()
	var daynote = this.getDaynoteText()

    //dateController
	var dateControl = (
		<div className="bookingView_header">
			<ToolBar wrapping="nowrap" className="bookingView_header_toolBar">
				<Datepicker
					id="booking.booking.date"
					value={this.state.date}
					fieldName={'date'}
					format="ddd Do MMM"
					canToggle={true}
					onChange={this.onChange}
				/>

				{allLoaded ?
					<ui.Select
						key="booking.timePeriods"
						id="booking.timePeriods"
						items={this.state.timePeriodStatisticsResult.value}
						dataValueField="name"
						dataTextField="text"
						value={this.state.selectedTimePeriodStatistics}
                        fieldName="timePeriod"
                        // handleTimePeriodChange, render-filter?
						onChange={this.handleTimePeriodChange}
					/>
				    : null}

                {allLoaded && hasValidDaynote ? 
                    <ui.Icon
						preset="daynote"
						onClick={this.handleToggleDaynote}
						className="bookingView_header_daynote_icon"
					/>
				    : null}
			</ToolBar>

            {allLoaded && hasValidDaynote && this.state.daynoteExpanded ?
                <div className="bookingView_header_daynote_text"> {daynote} </div>
			    : null}
		</div>
	);

	return (
		<ui.Wrap className="bookingView">
			<ColumnLayout
				activeColumn={this.state.activeColumn}
				onActiveChanged={this.handleActiveColumnChange}
			>
				<Column
                // Unit Controller
					sidebar={true}
					canGoLeft={true}
					name="PickUnit"
					title={this.t('booking.unit.pickUnit')}
				>
					<VerticalLayout>
						{allLoaded ?
							<ui.TableView
								id={'units'}
								items={this.state.unitsResult.value}
                                itemKey={'id'}
                                // handleSelectUnit. Probably dictates the group in the beginning of function and reruns this function upon selection,
								// (referencing: var groups = this.state.selectedUnit.articleGroups)
								onSelectItem={this.handleSelectUnit}
								selectedItem={null}
								searchBar={false}
								columns={[
									{
										name: 'units',
										type: 'header',
										title: this.t('booking.booking.time'),
										value: 'name',
									},
								]}
								clickWillTransition={true}
							/>
						    : <ui.FlashBox type="loading" />
						}
					</VerticalLayout>
				</Column>

				<Column
                // Group-render
					sidebar={true}
					canGoLeft={true}
					name="PickBooking"
					title={this.getUnitTitle(this.state.selectedUnit)}
					actionComponents={actions}
					controls={dateControl}
				>
					
					<VerticalLayout>
                        { // Where the evaluated groups are inserted
                        allLoaded ? groupsToRender
                            : <ui.FlashBox type="loading" />}
					</VerticalLayout>
				</Column>
			</ColumnLayout>

            { // Only displayed if in New-booking mode?
            this.state.newBooking && hasEditPermissions ?
				<EditBooking
					createMode       = {true}
					unit             = {this.state.selectedUnit}
					initialStartDate = {Moment(this.state.startFilter)}
					closeModal       = {this.handleCloseModal}
				/>
			: false}

            { // Only displayed if in Edit-booking mode?
            this.state.editBooking ?
				<EditBooking
					createMode={false}
					unit={this.state.selectedUnit}
					booking={this.state.editBooking}
					closeModal={this.handleCloseModal}
				/>
			: false}
		</ui.Wrap>
	)
}